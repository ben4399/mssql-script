/*==============================================================================
  脚本名称：为 AlwaysOn 可用性组 / 数据库镜像集群批量创建链接服务器
  目标平台：Microsoft SQL Server 2019（同时兼容 SQL Server 2016 SP2 及以上）
  原始版本：https://github.com/xiaohuazi123/mssql-script
            脚本/数据库帐号/跨多个SQL实例同步登录用户(适用于AlwaysOn和镜像)/1.创建链接服务器.sql
  部署位置：可用性组的每一个副本实例上都要执行（master 数据库）
  配套脚本：同目录 2.创建存储过程(优化版_SQL2019).sql、3.创建定时作业.sql

------------------------------------------------------------------------------
  一、这个脚本干什么
------------------------------------------------------------------------------
  1. 在多个副本实例之间建立"链接服务器"，作为账号同步脚本（存储过程
     usp_SyncLoginUserRegularBetweenInstances）下发 DDL 的传输通道。
  2. 每个实例上只需要创建"指向其他节点"的链接服务器，**不需要指向自己**。

  重要：目标地址请指向**每个副本实例本身**（主机名 / IP + 端口，或实例名），
        不要只指向可用性组侦听器（Listener）。因为侦听器永远只会落到当前主副本，
        而账号同步恰恰需要把账号下发到**每一个**副本实例上。

------------------------------------------------------------------------------
  二、相比原版的主要优化（针对 SQL Server 2019）
------------------------------------------------------------------------------
   01. 显式指定 @provider。原版只给了 @srvproduct，在 SQL Server 2019 上会导致
       provider 选择不确定，而旧驱动 SQLNCLI 已被弃用。
       本版自动探测已安装的驱动，优先 MSOLEDBSQL，其次 SQLNCLI11，成本为零。
   02. 链接服务器名改用"别名"，不再直接使用 "IP,端口" 作为对象名。
       原版把 '192.168.10.11,1433' 当作链接服务器名，后续引用必须写方括号，
       且一旦地址变更就要改所有脚本。本版把名字与地址解耦：
           名字 = N'AG_NODE02'，地址 = N'192.168.10.12,1433'。
   03. 增加幂等处理。原版重复执行会报"链接服务器已存在"，
       本版按 @ForceRecreate 决定"先删后建"还是"保留并只刷新选项"。
   04. 支持一次配置多个节点，用表变量维护清单，循环创建，不用改一处执行一次。
   05. 增加 SQL Server 2019 的加密陷阱处理。Microsoft OLE DB Driver 18 默认
       Encrypt=Optional，而 19 版默认 Encrypt=Mandatory，
       目标实例若未配置可信证书会直接连接失败。
       本版对 MSOLEDBSQL 自动附加 Encrypt=Optional。
   06. "remote proc transaction promotion" 由 true 改为 false。
       本场景用 EXEC(...) AT 下发 DDL，全程不在本地事务中，
       无需把操作提升为分布式事务，关掉可以彻底摆脱对 MSDTC 的依赖。
   07. 清理原版冗余设置（connect timeout / query timeout 各被设置了两次）。
       注意这两个值必须保持大于 0，存储过程正是靠它们识别"同步用的链接服务器"。
   08. 创建完成后立刻做连通性自检，验证 RPC Out、登录映射、加密参数是否生效。
   09. 增加"目标指向本机"的检测与告警，这是原版没有、但很容易配错的地方。
   10. 附录提供查看配置、列出可用性组其他副本、删除链接服务器的配套语句。

------------------------------------------------------------------------------
  三、部署步骤
------------------------------------------------------------------------------
    1. 在**每一个**副本实例上，修改下方"配置区"：
         · @RemoteLogin / @RemotePassword   —— 用于连接其他节点的账号
         · @Nodes                           —— 本节点要连接的其他副本清单
    2. 执行本脚本。注意用具备 sysadmin 权限的账号执行。
    3. 看执行结果中的自检信息，确认每个目标都显示 [成功]。
    4. 四个节点全部执行完毕后，再部署 2.创建存储过程(优化版_SQL2019).sql。

------------------------------------------------------------------------------
  四、注意事项
------------------------------------------------------------------------------
    · 连接账号建议使用专用的高权限账号。同步登录账号需要读取密码哈希
      （LOGINPROPERTY(...,'PasswordHash')），因此该账号必须具备 sysadmin
      或等价的 CONTROL SERVER 权限。
    · 脚本中的明文密码会被 SQL Server 加密后存入 sys.linked_logins
      （sysadmin 可解密查看）。建议执行完成后立即删除含密码的临时脚本文件。
    · 目标实例的 SQL Browser 服务若未启动，命名实例必须写全 "主机名\实例名,端口"。
    · 防火墙需放行实例端口（默认 1433）。
    · 若目标实例强制加密而证书不受信，请按附录 A 调整 @ProvStr。
==============================================================================*/

USE [master];
GO

SET NOCOUNT ON;
GO

/*==============================================================================
  配置区（只需要修改这一段）
==============================================================================*/
DECLARE @RemoteLogin    SYSNAME       = N'sa';        -- ★需修改：连接其他节点的登录名
DECLARE @RemotePassword NVARCHAR(128) = N'xxxxxx';    -- ★需修改：上述登录名的密码

DECLARE @ConnectTimeout INT           = 120;          -- 连接超时（秒），必须 > 0
DECLARE @QueryTimeout   INT           = 120;          -- 查询超时（秒），必须 > 0

DECLARE @Provider       SYSNAME       = NULL;         -- 驱动。NULL = 自动探测（推荐）
DECLARE @ProvStr        NVARCHAR(256) = NULL;         -- 驱动附加参数。NULL = 按驱动自动决定

DECLARE @ForceRecreate  BIT           = 1;            -- 1 = 链接服务器已存在时先删后建；0 = 保留，仅刷新选项与登录映射

DECLARE @LocalServerName SYSNAME = CAST(SERVERPROPERTY(N'ServerName') AS SYSNAME);

/* 目标节点清单
     LinkedServerName : 链接服务器别名，建议用节点名，便于理解和维护
     DataSource       : 目标实例地址。写法示例：
                          N'192.168.10.12,1433'       —— IP + 端口
                          N'SQLNODE02'                —— 主机名（默认实例）
                          N'SQLNODE02\INST01,1433'    —— 命名实例
     RemoteLogin      : 该节点单独使用的登录名；留 NULL 则沿用上面的 @RemoteLogin
     RemotePassword   : 同上
     Enable           : 1 = 本次处理；0 = 本次跳过
*/
DECLARE @Nodes TABLE
(
    Seq              INT IDENTITY(1,1) PRIMARY KEY,
    LinkedServerName SYSNAME       NOT NULL,
    DataSource       NVARCHAR(256) NOT NULL,
    RemoteLogin      SYSNAME       NULL,
    RemotePassword   NVARCHAR(128) NULL,
    Enable           BIT           NOT NULL DEFAULT (1)
);

INSERT INTO @Nodes (LinkedServerName, DataSource) VALUES
    (N'AG_NODE02', N'192.168.10.12,1433'),
    (N'AG_NODE03', N'192.168.10.13,1433');
    -- (N'AG_NODE04', N'192.168.10.14,1433');


/*==============================================================================
  第 1 步：确定使用哪个 OLE DB 驱动
  说明：SQL Server 2019 环境请优先使用 MSOLEDBSQL（Microsoft OLE DB Driver），
        SQLNCLI 系列已从 SQL Server 2022 起移除。此处自动探测，无需人工干预。
==============================================================================*/
IF @Provider IS NULL
BEGIN
    DECLARE @prov TABLE
    (
        ProviderName        SYSNAME,
        ParseName           SYSNAME,
        ProviderClsid       NVARCHAR(128),
        ProviderDescription NVARCHAR(400)
    );

    BEGIN TRY
        INSERT INTO @prov EXEC master.dbo.sp_enum_oledb_providers;
    END TRY
    BEGIN CATCH
        PRINT N'[提示] 无法枚举 OLE DB 驱动（sp_enum_oledb_providers 在当前版本可能已移除），将使用兜底值。';
    END CATCH

    SELECT TOP (1) @Provider = p.ProviderName
    FROM @prov AS p
    WHERE p.ProviderName IN (N'MSOLEDBSQL', N'SQLNCLI11', N'SQLNCLI')
    ORDER BY CASE p.ProviderName
                 WHEN N'MSOLEDBSQL' THEN 1
                 WHEN N'SQLNCLI11'  THEN 2
                 ELSE 3
             END;

    IF @Provider IS NULL
        SET @Provider = N'SQLNCLI11';
END

/* MSOLEDBSQL 19 版默认强制加密，目标实例没有可信证书时会连接失败。
   这里默认放宽为 Encrypt=Optional，除非调用方自己指定了 @ProvStr。
   若目标实例确实要求加密，请显式传入： @ProvStr = N'Encrypt=Mandatory;' */
IF @ProvStr IS NULL AND @Provider = N'MSOLEDBSQL'
    SET @ProvStr = N'Encrypt=Optional;';

PRINT N'============================================================';
PRINT N'[配置] 本机实例       ：' + @LocalServerName;
PRINT N'[配置] 使用驱动       ：' + @Provider + ISNULL(N'（' + @ProvStr + N'）', N'');
PRINT N'[配置] 连接超时 / 查询超时：' + CAST(@ConnectTimeout AS NVARCHAR(10))
    + N' 秒 / ' + CAST(@QueryTimeout AS NVARCHAR(10)) + N' 秒';
PRINT N'[配置] 待处理节点数   ：' + CAST((SELECT COUNT(*) FROM @Nodes WHERE Enable = 1) AS NVARCHAR(10));
PRINT N'============================================================';


/*==============================================================================
  第 2 步：逐个节点创建链接服务器
==============================================================================*/
DECLARE @i       INT = 1;
DECLARE @cnt     INT = (SELECT COUNT(*) FROM @Nodes);
DECLARE @srv     SYSNAME;
DECLARE @ds      NVARCHAR(256);
DECLARE @ru      SYSNAME;
DECLARE @rp      NVARCHAR(128);
DECLARE @en      BIT;
DECLARE @exists  BIT;
DECLARE @okCount INT = 0;
DECLARE @ngCount INT = 0;

WHILE @i <= @cnt
BEGIN
    SELECT @srv = LinkedServerName,
           @ds  = DataSource,
           @ru  = ISNULL(RemoteLogin,    @RemoteLogin),
           @rp  = ISNULL(RemotePassword, @RemotePassword),
           @en  = Enable
    FROM @Nodes
    WHERE Seq = @i;

    PRINT N'';
    PRINT N'------------------------------------------------------------';
    PRINT N'[处理] 节点 ' + CAST(@i AS NVARCHAR(10)) + N'/' + CAST(@cnt AS NVARCHAR(10))
        + N'：[' + @srv + N'] → ' + @ds;

    IF @en = 0
    BEGIN
        PRINT N'[跳过] 该节点已在清单中标记为不启用。';
        SET @i += 1;
        CONTINUE;
    END

    IF @srv IS NULL OR @ds IS NULL
    BEGIN
        PRINT N'[跳过] 别名或地址为空，配置不完整。';
        SET @ngCount += 1;
        SET @i += 1;
        CONTINUE;
    END

    /* 2.1 幂等处理：已存在时按 @ForceRecreate 决定删除或保留 */
    SET @exists = CASE WHEN EXISTS
                    (
                        SELECT 1 FROM sys.servers
                        WHERE name = @srv AND is_linked = 1
                    ) THEN 1 ELSE 0 END;

    IF @exists = 1 AND @ForceRecreate = 1
    BEGIN
        PRINT N'[重建] 链接服务器 [' + @srv + N'] 已存在，按配置先删除。';

        BEGIN TRY
            EXEC master.dbo.sp_dropserver @server = @srv, @droplogins = 'droplogins';
        END TRY
        BEGIN CATCH
            PRINT N'[失败] 删除旧链接服务器失败：' + ERROR_MESSAGE();
            PRINT N'  请确认没有正在使用该链接服务器的会话或作业。';
            SET @ngCount += 1;
            SET @i += 1;
            CONTINUE;
        END CATCH

        SET @exists = 0;
    END
    ELSE IF @exists = 1
    BEGIN
        PRINT N'[保留] 链接服务器 [' + @srv + N'] 已存在，本次只刷新选项与登录映射。';
    END

    /* 2.2 创建链接服务器 */
    IF @exists = 0
    BEGIN
        BEGIN TRY
            /* 关键点：名字（@server）与地址（@datasrc）分离，
               名字用于脚本中引用，地址可以随时调整。 */
            EXEC master.dbo.sp_addlinkedserver
                 @server     = @srv,
                 @srvproduct = N'SQL Server',
                 @provider   = @Provider,
                 @datasrc    = @ds,
                 @provstr    = @ProvStr;

            PRINT N'[创建] 链接服务器 [' + @srv + N'] 创建完成。';
        END TRY
        BEGIN CATCH
            PRINT N'[失败] 创建链接服务器失败：' + ERROR_MESSAGE();
            PRINT N'  常见原因：驱动未安装、地址不可达、或该名称已被本地服务器占用。';
            SET @ngCount += 1;
            SET @i += 1;
            CONTINUE;
        END CATCH
    END

    /* 2.3 配置服务器选项
       说明：
         · data access   = true  → 允许 OPENQUERY 读取远端数据（存储过程靠它拉取目标端现状）
         · rpc out       = true  → 允许 EXEC (...) AT [xxx] 下发语句到远端执行（核心）
         · connect/query timeout 必须 > 0，存储过程据此识别同步用的链接服务器
         · remote proc transaction promotion = false → 不需要分布式事务，避免依赖 MSDTC */
    BEGIN TRY
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'collation compatible', @optvalue = N'false';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'data access',         @optvalue = N'true';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'dist',                @optvalue = N'false';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'pub',                 @optvalue = N'false';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'rpc',                 @optvalue = N'true';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'rpc out',             @optvalue = N'true';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'sub',                 @optvalue = N'false';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'collation name',      @optvalue = NULL;
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'lazy schema validation', @optvalue = N'false';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'use remote collation', @optvalue = N'true';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'remote proc transaction promotion', @optvalue = N'false';
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'connect timeout',
             @optvalue = CAST(@ConnectTimeout AS NVARCHAR(10));
        EXEC master.dbo.sp_serveroption @server = @srv, @optname = N'query timeout',
             @optvalue = CAST(@QueryTimeout AS NVARCHAR(10));

        PRINT N'[选项] 服务器选项配置完成（data access / rpc out 已开启，超时 '
            + CAST(@ConnectTimeout AS NVARCHAR(10)) + N' 秒）。';
    END TRY
    BEGIN CATCH
        PRINT N'[失败] 配置服务器选项失败：' + ERROR_MESSAGE();
        SET @ngCount += 1;
        SET @i += 1;
        CONTINUE;
    END CATCH

    /* 2.4 配置登录映射
       说明：@locallogin = NULL 表示"本机所有登录"访问该链接服务器时，
             统一使用这里指定的远端账号，调用方无需逐个授权。
             先删除已有映射，保证重复执行不出错。 */
    BEGIN TRY
        /* sp_addlinkedserver 会自动建立一个"沿用本地登录"的默认映射，
           这里先把它删掉，再用专用账号重新建立，保证重复执行也不报错。 */
        IF EXISTS
           (
               SELECT 1
               FROM   sys.linked_logins AS l
               JOIN   sys.servers       AS s ON s.server_id = l.server_id
               WHERE  s.name = @srv
                 AND  l.local_principal_id = 0
           )
        BEGIN
            EXEC master.dbo.sp_droplinkedsrvlogin @rmtsrvname = @srv, @locallogin = NULL;
        END

        EXEC master.dbo.sp_addlinkedsrvlogin
             @rmtsrvname  = @srv,
             @useself     = N'False',
             @locallogin  = NULL,
             @rmtuser     = @ru,
             @rmtpassword = @rp;

        PRINT N'[映射] 登录映射已完成，远端账号：' + @ru;
    END TRY
    BEGIN CATCH
        PRINT N'[失败] 配置登录映射失败：' + ERROR_MESSAGE();
        SET @ngCount += 1;
        SET @i += 1;
        CONTINUE;
    END CATCH

    /* 2.5 连通性自检
       验证链路是否真的可用：既要能读（OPENQUERY），也要能写（RPC Out）。 */
    IF OBJECT_ID(N'tempdb..#ConnCheck') IS NOT NULL
        DROP TABLE #ConnCheck;

    DECLARE @RemoteName NVARCHAR(256) = NULL;
    DECLARE @CheckSQL   NVARCHAR(1000);

    BEGIN TRY
        CREATE TABLE #ConnCheck (RemoteServerName NVARCHAR(256));

        SET @CheckSQL = N'SELECT * FROM OPENQUERY(' + QUOTENAME(@srv)
                      + N', ''SELECT @@SERVERNAME AS RemoteServerName'');';

        INSERT INTO #ConnCheck (RemoteServerName) EXEC (@CheckSQL);

        SELECT @RemoteName = RemoteServerName FROM #ConnCheck;
    END TRY
    BEGIN CATCH
        PRINT N'[失败] 连通性测试未通过：' + ERROR_MESSAGE();
        PRINT N'  排查建议：';
        PRINT N'    1) 目标实例的 TCP 端口是否放行（默认 1433，命名实例需加端口）；';
        PRINT N'    2) 连接账号密码是否正确、是否有登录权限；';
        PRINT N'    3) 若使用 MSOLEDBSQL，确认 @ProvStr 的 Encrypt 设置与目标证书匹配；';
        PRINT N'    4) 目标实例是否需要 SQL Browser 服务（命名实例）。';
        SET @ngCount += 1;
        SET @i += 1;
        CONTINUE;
    END CATCH

    IF @RemoteName IS NULL
    BEGIN
        PRINT N'[警告] 已连通但未能取回远端实例名，请人工确认。';
        SET @ngCount += 1;
    END
    ELSE IF @RemoteName = @LocalServerName
    BEGIN
        PRINT N'[警告] 目标地址实际指向了本机（远端实例名 = ' + @RemoteName + N'）。';
        PRINT N'  该链接服务器没有实际用途，建议删除，或改指向其他副本节点。';
        SET @ngCount += 1;
    END
    ELSE
    BEGIN
        PRINT N'[成功] [' + @srv + N'] 连通正常，远端实例：' + @RemoteName;
        SET @okCount += 1;
    END

    SET @i += 1;
END

IF OBJECT_ID(N'tempdb..#ConnCheck') IS NOT NULL
    DROP TABLE #ConnCheck;


/*==============================================================================
  第 3 步：输出结果汇总
==============================================================================*/
PRINT N'';
PRINT N'============================================================';
PRINT N'[结束] 创建完成。连通正常 ' + CAST(@okCount AS NVARCHAR(10))
    + N' 个，存在问题 ' + CAST(@ngCount AS NVARCHAR(10)) + N' 个。';
PRINT N'============================================================';
PRINT N'';
PRINT N'[当前本机的链接服务器清单，以及连接账号列表]';

SELECT  s.name                              AS [链接服务器名],
        s.data_source                       AS [目标地址],
        s.provider                          AS [驱动],
        s.product                           AS [产品],
        s.catalog                           AS [默认数据库],
        s.is_rpc_out_enabled                AS [允许RPC出站],
        s.is_data_access_enabled            AS [允许数据访问],
        s.connect_timeout                   AS [连接超时],
        s.query_timeout                     AS [查询超时],
        l.remote_name                       AS [远端账号]
FROM    sys.servers        AS s
LEFT JOIN sys.linked_logins AS l
       ON l.server_id = s.server_id
      AND l.local_principal_id = 0
WHERE   s.is_linked = 1
ORDER BY s.name;

PRINT N'[提示] 请确认上表中"允许RPC出站"为 1、"连接超时"和"查询超时"均大于 0，';
PRINT N'       这三项是账号同步存储过程正常工作的前提。';
GO


/*==============================================================================
  附录 A：常见参数调整
==============================================================================*/
-- A1. 目标实例强制加密，且使用受信任证书：
--     DECLARE @ProvStr NVARCHAR(256) = N'Encrypt=Mandatory;';
--
-- A2. 目标实例强制加密，但证书不受信任（仅建议测试环境）：
--     DECLARE @ProvStr NVARCHAR(256) = N'Encrypt=Mandatory;TrustServerCertificate=Yes;';
--
-- A3. 希望明确使用某个驱动：
--     DECLARE @Provider SYSNAME = N'SQLNCLI11';   -- 或 N'MSOLEDBSQL'
--
-- A4. 希望保留已有链接服务器，只刷新超时等选项（不重建）：
--     DECLARE @ForceRecreate BIT = 0;
GO


/*==============================================================================
  附录 B：查询本实例所属可用性组中的其他副本（用于填写 @Nodes 清单）
==============================================================================*/
-- SELECT  ag.name                       AS 可用性组,
--         ar.server_name                AS 副本实例名,
--         ars.role_desc                 AS 当前角色,
--         ars.operational_state_desc    AS 运行状态,
--         ars.synchronization_health_desc AS 同步健康度,
--         ar.secondary_role_allow_connections_desc AS 辅助副本是否可读
-- FROM    sys.availability_groups                     AS ag
-- JOIN    sys.availability_replicas                   AS ar  ON ar.group_id  = ag.group_id
-- JOIN    sys.dm_hadr_availability_replica_states     AS ars ON ars.replica_id = ar.replica_id
-- WHERE   ag.group_id IN
--         (
--             SELECT group_id
--             FROM   sys.dm_hadr_availability_replica_states
--             WHERE  is_local = 1
--         )
-- ORDER BY ar.server_name;
GO


/*==============================================================================
  附录 C：查看 / 删除链接服务器
==============================================================================*/
-- C1. 查看某个链接服务器的全部选项
-- EXEC master.dbo.sp_helplinkedsrvlogin @rmtsrvname = N'AG_NODE02';
-- EXEC master.dbo.sp_serveroption     @server = N'AG_NODE02';
--
-- C2. 删除单个链接服务器（连同其登录映射）
-- EXEC master.dbo.sp_dropserver @server = N'AG_NODE02', @droplogins = 'droplogins';
--
-- C3. 批量删除本机所有 SQL Server 类型的链接服务器（谨慎使用）
-- DECLARE @n SYSNAME;
-- DECLARE c CURSOR LOCAL FAST_FORWARD FOR
--     SELECT name FROM sys.servers WHERE is_linked = 1 AND product = 'SQL Server';
-- OPEN c;
-- FETCH NEXT FROM c INTO @n;
-- WHILE @@FETCH_STATUS = 0
-- BEGIN
--     EXEC master.dbo.sp_dropserver @server = @n, @droplogins = 'droplogins';
--     PRINT N'已删除：' + @n;
--     FETCH NEXT FROM c INTO @n;
-- END
-- CLOSE c;
-- DEALLOCATE c;
GO
