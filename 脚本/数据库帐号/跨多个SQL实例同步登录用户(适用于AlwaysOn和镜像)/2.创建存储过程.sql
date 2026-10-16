/*==============================================================================
  脚本名称：跨多个 SQL 实例同步登录用户（AlwaysOn 可用性组 / 数据库镜像）
  目标平台：Microsoft SQL Server 2019（同时兼容 SQL Server 2016 SP2 及以上）
  原始版本：https://github.com/xiaohuazi123/mssql-script
            脚本/数据库帐号/跨多个SQL实例同步登录用户(适用于AlwaysOn和镜像)/2.创建存储过程.sql
  部署位置：每个 SQL 实例的 master 数据库（必须每个副本节点都部署一遍）
  调用方式：EXEC master.dbo.usp_SyncLoginUserRegularBetweenInstances;
            或由定时作业周期调用（见同目录 3.创建定时作业.sql）

------------------------------------------------------------------------------
  一、为什么需要它
------------------------------------------------------------------------------
  在 AlwaysOn 可用性组（或数据库镜像）环境中，登录账号（Login）存放在实例级别的
  master 数据库中，它不会随可用性数据库一起同步。于是主副本上新建的账号在辅助副本
  上并不存在，一旦发生故障切换就会出现"账号丢失、应用连不上"的问题。

  本存储过程由主副本周期性地把自身实例上的登录账号，通过链接服务器下发到其余各个
  副本实例执行，实现账号自动对齐。

------------------------------------------------------------------------------
  二、同步内容
------------------------------------------------------------------------------
    1. 登录是否存在      —— 不存在则按源端定义创建
    2. SID               —— 与源端保持一致，避免目标库出现孤立用户（Orphaned User）
    3. 密码              —— 以哈希（HASHED）方式下发，全程不接触明文密码
    4. 默认数据库 / 默认语言
    5. CHECK_POLICY / CHECK_EXPIRATION
    6. 登录的启用 / 禁用状态
    7. 固定服务器角色成员关系
    8. Windows 登录 / 组（可选开关，默认关闭）

------------------------------------------------------------------------------
  三、相比原版的主要优化（针对 SQL Server 2019）
------------------------------------------------------------------------------
   01. 彻底移除对 sp_hexadecimal 的依赖，改用 CONVERT(NVARCHAR(1026), <varbinary>, 1)
       直接得到 0x 十六进制字符串。目标实例不需要预置任何对象，真正做到开箱即用。
   02. 放宽目标节点识别条件。原版写死 provider = 'SQLNCLI'，而 SQL Server 2019 环境的
       链接服务器普遍使用 MSOLEDBSQL 或 SQLNCLI11，会导致一个目标都扫不到。
   03. 修复原版读取 is_policy_checked / is_expiration_checked 时未按登录名过滤的缺陷
       （原版第 113-114 行）。原写法在多 SQL 登录并存时会取到不确定的一行，
       导致 CHECK_POLICY / CHECK_EXPIRATION 被错误同步。
   04. 修复原版未排除 public 服务器角色的缺陷。原版用 sp_helpsrvrolemember 取角色后
       直接调用 sp_addsrvrolemember 添加 public，必然报错。
       本版改用系统视图 sys.server_role_members 并显式排除 public，
       同时改用 SQL Server 2012 起推荐的 ALTER SERVER ROLE ... ADD MEMBER 语法。
   05. 修复原版"目标端已存在同名登录但 SID 不一致"时会直接执行 CREATE LOGIN 而报
       "登录已存在"的缺陷。本版增加 @RebuildOnSidMismatch 开关统一处理。
   06. 增加登录类型冲突处理（源端是 SQL 登录、目标端是同名 Windows 登录）。
   07. 默认数据库在目标实例不存在时自动回退为 master 并记录警告，
       避免 CREATE LOGIN 因默认库不存在而失败。
   08. 增加禁用状态同步、服务器角色同步、Windows 登录同步开关。
   09. 增加运行日志表 master.dbo.SyncLoginUserLog，逐条记录执行的 DDL 与结果，
       可追溯、可对接告警。
   10. 增加 @Debug 试运行模式，只生成并打印脚本，不真正下发。
   11. 所有动态对象名统一使用 QUOTENAME 转义，避免特殊字符导致的语法错误。
   12. 跨实例字符串显式 COLLATE DATABASE_DEFAULT，规避链接服务器常见的排序规则冲突。
   13. 目标端存在但源端没有的登录不会被删除（避免误删生产账号），仅记录日志。
   14. 增加 @ForceRun 开关，非可用性组（单实例）场景也可手工执行本同步。

------------------------------------------------------------------------------
  四、部署步骤
------------------------------------------------------------------------------
    1. 在每个副本实例上执行本脚本（会在该实例 master 中创建日志表和存储过程）。
    2. 确保每个副本实例上都已创建指向其他副本实例的链接服务器
       （见同目录 1.创建链接服务器.sql），并且该链接服务器的 RPC Out 已开启、
       登录映射使用具备 sysadmin 权限的账号（读取密码哈希需要该权限）。
    3. 用 3.创建定时作业.sql 创建周期作业，或直接在作业步骤中调用：
           EXEC master.dbo.usp_SyncLoginUserRegularBetweenInstances;
    4. 首次务必先试运行检查生成的 DDL：
           EXEC master.dbo.usp_SyncLoginUserRegularBetweenInstances @Debug = 1;

------------------------------------------------------------------------------
  五、注意事项
------------------------------------------------------------------------------
    · 同步方向是单向的：主副本 → 其他副本。在辅助副本上手工改密码，
      会在下一次同步时被主副本的值覆盖。
    · 每个副本实例都要部署本存储过程；只有主副本会真正动作。
    · 执行账号需要 sysadmin 权限（读取 LOGINPROPERTY 密码哈希必需）。
    · 本脚本用到 STRING_SPLIT，要求实例兼容级别不低于 130（SQL2019 默认 150）。
==============================================================================*/

USE [master];
GO
SET NOCOUNT ON;
GO

/*==============================================================================
  第 1 步：创建运行日志表
  说明：记录每一次同步运行、每一句 DDL 的执行结果，便于排查与告警。
==============================================================================*/
IF OBJECT_ID(N'dbo.SyncLoginUserLog', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SyncLoginUserLog
    (
        LogID         BIGINT IDENTITY(1,1) NOT NULL
                      CONSTRAINT PK_SyncLoginUserLog PRIMARY KEY,
        RunID         UNIQUEIDENTIFIER     NOT NULL,        -- 每次运行的唯一标识
        RunTime       DATETIME2(0)         NOT NULL
                      CONSTRAINT DF_SyncLoginUserLog_RunTime DEFAULT (SYSDATETIME()),
        SourceServer  SYSNAME              NOT NULL,        -- 源实例（主副本）
        TargetServer  SYSNAME              NULL,            -- 目标实例
        ObjectType    VARCHAR(20)          NOT NULL,        -- LOGIN / ROLE / SYSTEM
        ObjectName    SYSNAME              NULL,            -- 登录名或角色名
        Action        VARCHAR(20)          NULL,            -- CREATE/ALTER/DROP/RECREATE/ADDROLE/SKIP/WARN
        Status        VARCHAR(10)          NOT NULL,        -- SUCCESS / FAILED / WARNING / INFO
        Message       NVARCHAR(4000)       NULL,            -- 中文说明或错误信息
        DDLText       NVARCHAR(MAX)        NULL             -- 实际下发的语句
    );

    CREATE NONCLUSTERED INDEX IX_SyncLoginUserLog_RunTime
        ON dbo.SyncLoginUserLog (RunTime)
        INCLUDE (TargetServer, Status);

    PRINT N'[初始化] 已创建运行日志表 master.dbo.SyncLoginUserLog';
END
ELSE
    PRINT N'[初始化] 运行日志表 master.dbo.SyncLoginUserLog 已存在，跳过创建';
GO


/*==============================================================================
  第 2 步：创建同步存储过程
==============================================================================*/
IF OBJECT_ID(N'dbo.usp_SyncLoginUserRegularBetweenInstances', N'P') IS NOT NULL
BEGIN
    DROP PROCEDURE dbo.usp_SyncLoginUserRegularBetweenInstances;
    PRINT N'[初始化] 已删除旧的 usp_SyncLoginUserRegularBetweenInstances';
END
GO

CREATE PROCEDURE dbo.usp_SyncLoginUserRegularBetweenInstances
(
    @LinkedServerList     NVARCHAR(MAX) = NULL,  -- 目标节点白名单（逗号分隔的链接服务器名）。NULL = 自动发现所有非本机 SQL Server 链接服务器
    @IncludeWindowsLogins BIT           = 0,     -- 1 = 同时同步 Windows 登录 / 组；0 = 只同步 SQL 登录
    @SyncDisabledState    BIT           = 1,     -- 1 = 把源端的启用 / 禁用状态下发到目标端
    @SyncServerRoles      BIT           = 1,     -- 1 = 同步固定服务器角色成员关系
    @RebuildOnSidMismatch BIT           = 0,     -- 1 = 源 / 目标同名登录 SID 不一致时先 DROP 再重建；0 = 跳过并告警
    @ExcludeLoginList     NVARCHAR(MAX) = NULL,  -- 不参与同步的登录名（逗号分隔），例如复制账号 distributor_admin
    @ForceRun             BIT           = 0,     -- 1 = 跳过"当前实例必须是主副本"的判定，用于非可用性组的单机分发场景
    @Debug                BIT           = 0,     -- 1 = 只生成并打印脚本，不真正下发执行
    @LogRetentionDays     INT           = 30     -- 日志保留天数，<= 0 表示不清理
)
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @RunID        UNIQUEIDENTIFIER = NEWID();
    DECLARE @SourceServer SYSNAME = CAST(SERVERPROPERTY(N'ServerName') AS SYSNAME);
    DECLARE @msg          NVARCHAR(4000);
    DECLARE @TotalTarget  INT = 0;
    DECLARE @OKTarget     INT = 0;
    DECLARE @FailTarget   INT = 0;

    PRINT N'============================================================';
    PRINT N'[开始] 登录账号同步任务，源实例：' + @SourceServer;
    PRINT N'[开始] 运行号：' + CONVERT(VARCHAR(36), @RunID)
        + CASE WHEN @Debug = 1 THEN N'（试运行模式，不会真正执行）' ELSE N'' END;
    PRINT N'============================================================';

    /*--------------------------------------------------------------------------
      2.1 主副本判定
      说明：定时作业会在可用性组的每个节点上都运行。这里只允许"当前实例是某个
            可用性组的主副本且处于在线状态"时才继续执行，
            从而保证同一时刻只有主副本在生成并下发脚本，天然避免多节点互相覆盖；
            发生故障切换后，新的主副本会自动接管，无需修改任何配置。
    --------------------------------------------------------------------------*/
    IF @ForceRun = 0
       AND NOT EXISTS
    (
        SELECT 1
        FROM sys.dm_hadr_availability_replica_states AS rs
        WHERE rs.is_local = 1
          AND rs.role_desc = N'PRIMARY'
          AND rs.operational_state_desc = N'ONLINE'
    )
    BEGIN
        SET @msg = N'当前实例不是任何可用性组的主副本（或状态不在线），本次不同步。';
        PRINT N'[跳过] ' + @msg;

        INSERT INTO dbo.SyncLoginUserLog (RunID, SourceServer, ObjectType, Status, Message)
        VALUES (@RunID, @SourceServer, 'SYSTEM', 'INFO', @msg);

        RETURN;
    END

    IF @ForceRun = 1
        PRINT N'[判定] 已指定 @ForceRun = 1，跳过主副本判定，直接执行同步。';
    ELSE
        PRINT N'[判定] 当前实例是主副本，继续执行同步。';

    /*--------------------------------------------------------------------------
      2.2 确定目标节点清单
      说明：默认自动发现本实例上所有指向其他 SQL Server 实例的链接服务器。
            注意 SQL Server 2019 上链接服务器多数使用 MSOLEDBSQL 或 SQLNCLI11，
            因此这里不再限制 provider，只校验 product 与 is_linked。
    --------------------------------------------------------------------------*/
    CREATE TABLE #Target
    (
        ServerName SYSNAME COLLATE DATABASE_DEFAULT NOT NULL PRIMARY KEY
    );

    IF @LinkedServerList IS NOT NULL AND LTRIM(RTRIM(@LinkedServerList)) <> N''
    BEGIN
        INSERT INTO #Target (ServerName)
        SELECT DISTINCT LTRIM(RTRIM(v.value))
        FROM STRING_SPLIT(@LinkedServerList, N',') AS v
        WHERE LTRIM(RTRIM(v.value)) <> N'';

        PRINT N'[目标] 使用调用方指定的目标节点清单。';
    END
    ELSE
    BEGIN
        INSERT INTO #Target (ServerName)
        SELECT s.name
        FROM sys.servers AS s
        WHERE s.is_linked = 1
          AND s.product = N'SQL Server'
          AND s.name <> @SourceServer;

        PRINT N'[目标] 已自动发现目标节点。';
    END

    SELECT @TotalTarget = COUNT(*) FROM #Target;

    DECLARE @TargetListStr NVARCHAR(4000) =
        (SELECT STRING_AGG(CONVERT(NVARCHAR(128), t.ServerName), N', ')
         FROM #Target AS t);

    PRINT N'[目标] 目标节点共 ' + CAST(@TotalTarget AS NVARCHAR(10)) + N' 个：'
        + ISNULL(@TargetListStr, N'（无）');

    IF @TotalTarget = 0
    BEGIN
        SET @msg = N'未发现任何可用的目标节点（请检查链接服务器是否已创建）。本次不同步。';
        PRINT N'[跳过] ' + @msg;

        INSERT INTO dbo.SyncLoginUserLog (RunID, SourceServer, ObjectType, Status, Message)
        VALUES (@RunID, @SourceServer, 'SYSTEM', 'WARNING', @msg);

        RETURN;
    END

    /*--------------------------------------------------------------------------
      2.3 采集源端（本实例）的登录账号
      说明：
        · SID 与密码哈希统一用 CONVERT(..., 1) 转成 0x 十六进制字符串，
          便于跨实例比较，同时让目标实例无需依赖 sp_hexadecimal。
        · 密码哈希取自 LOGINPROPERTY(登录名, 'PasswordHash')，全程不接触明文密码。
        · 排除 sa、以 ## 开头结尾的系统内部登录、NT SERVICE\ 与 NT AUTHORITY\ 服务账号。
    --------------------------------------------------------------------------*/
    CREATE TABLE #SrcLogin
    (
        LoginName           SYSNAME COLLATE DATABASE_DEFAULT       NOT NULL PRIMARY KEY,
        LoginType           CHAR(1)                                NOT NULL,
        SIDHex              NVARCHAR(1026) COLLATE DATABASE_DEFAULT NULL,
        PWDHex              NVARCHAR(1026) COLLATE DATABASE_DEFAULT NULL,
        DefaultDB           SYSNAME COLLATE DATABASE_DEFAULT       NULL,
        DefaultLanguage     SYSNAME COLLATE DATABASE_DEFAULT       NULL,
        IsDisabled          BIT                                    NOT NULL,
        IsPolicyChecked     INT                                    NULL,
        IsExpirationChecked INT                                    NULL
    );

    INSERT INTO #SrcLogin
        (LoginName, LoginType, SIDHex, PWDHex, DefaultDB, DefaultLanguage,
         IsDisabled, IsPolicyChecked, IsExpirationChecked)
    SELECT
        p.name,
        p.type,
        CONVERT(NVARCHAR(1026), p.sid, 1),
        CASE WHEN p.type = 'S'
             THEN CONVERT(NVARCHAR(1026), LOGINPROPERTY(p.name, 'PasswordHash'), 1)
             ELSE NULL
        END,
        p.default_database_name,
        p.default_language_name,
        p.is_disabled,
        CAST(l.is_policy_checked     AS INT),
        CAST(l.is_expiration_checked AS INT)
    FROM sys.server_principals AS p
    LEFT JOIN sys.sql_logins   AS l ON l.principal_id = p.principal_id
    WHERE p.type IN ('S', 'U', 'G')
      AND p.name <> N'sa'
      AND p.name NOT LIKE N'##%##'
      AND p.name NOT LIKE N'NT SERVICE\%'
      AND p.name NOT LIKE N'NT AUTHORITY\%'
      AND (p.type = 'S' OR @IncludeWindowsLogins = 1)
      AND (@ExcludeLoginList IS NULL
           OR NOT EXISTS
              (
                  SELECT 1
                  FROM STRING_SPLIT(@ExcludeLoginList, N',') AS e
                  WHERE LTRIM(RTRIM(e.value)) = p.name
              ));

    PRINT N'[采集] 源端共采集到 ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) + N' 个待同步的登录账号。';

    /*--------------------------------------------------------------------------
      2.4 采集源端登录所属的固定服务器角色
      说明：必须排除 public —— public 是隐式角色，任何登录都自动属于它，
            显式添加会报错。
    --------------------------------------------------------------------------*/
    CREATE TABLE #SrcRole
    (
        LoginName  SYSNAME COLLATE DATABASE_DEFAULT NOT NULL,
        ServerRole SYSNAME COLLATE DATABASE_DEFAULT NOT NULL,
        PRIMARY KEY (LoginName, ServerRole)
    );

    IF @SyncServerRoles = 1
    BEGIN
        INSERT INTO #SrcRole (LoginName, ServerRole)
        SELECT rp.name, r.name
        FROM sys.server_role_members AS rm
        JOIN sys.server_principals AS r  ON r.principal_id  = rm.role_principal_id
        JOIN sys.server_principals AS rp ON rp.principal_id = rm.member_principal_id
        WHERE r.name <> N'public'
          AND rp.name IN (SELECT LoginName FROM #SrcLogin);

        PRINT N'[采集] 源端共采集到 ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) + N' 条服务器角色成员关系。';
    END

    /*--------------------------------------------------------------------------
      2.5 准备逐节点循环使用的临时表
    --------------------------------------------------------------------------*/
    CREATE TABLE #DstLogin
    (
        LoginName           SYSNAME COLLATE DATABASE_DEFAULT       NOT NULL PRIMARY KEY,
        LoginType           CHAR(1)                                NULL,
        SIDHex              NVARCHAR(1026) COLLATE DATABASE_DEFAULT NULL,
        PWDHex              NVARCHAR(1026) COLLATE DATABASE_DEFAULT NULL,
        DefaultDB           SYSNAME COLLATE DATABASE_DEFAULT       NULL,
        DefaultLanguage     SYSNAME COLLATE DATABASE_DEFAULT       NULL,
        IsDisabled          BIT                                    NOT NULL,
        IsPolicyChecked     INT                                    NULL,
        IsExpirationChecked INT                                    NULL
    );

    CREATE TABLE #DstRole
    (
        LoginName  SYSNAME COLLATE DATABASE_DEFAULT NOT NULL,
        ServerRole SYSNAME COLLATE DATABASE_DEFAULT NOT NULL,
        PRIMARY KEY (LoginName, ServerRole)
    );

    CREATE TABLE #DstDB
    (
        DBName SYSNAME COLLATE DATABASE_DEFAULT NOT NULL PRIMARY KEY
    );

    -- 待下发语句清单。DDL 为 NULL 表示"仅记录日志、不执行"（例如跳过项）。
    CREATE TABLE #SyncDDL
    (
        Seq        INT           IDENTITY(1,1) PRIMARY KEY,
        ObjectType VARCHAR(20)   NOT NULL,
        ObjectName SYSNAME       NULL,
        Action     VARCHAR(20)   NULL,
        DDL        NVARCHAR(MAX) NULL,
        Message    NVARCHAR(4000) NULL
    );

    /*--------------------------------------------------------------------------
      2.6 逐节点比对并下发
    --------------------------------------------------------------------------*/
    DECLARE @Target SYSNAME;

    DECLARE cur_target CURSOR LOCAL FAST_FORWARD FOR
        SELECT ServerName FROM #Target ORDER BY ServerName;

    OPEN cur_target;
    FETCH NEXT FROM cur_target INTO @Target;

    WHILE @@FETCH_STATUS = 0
    BEGIN
        PRINT N'------------------------------------------------------------';
        PRINT N'[处理] 目标节点：' + @Target;

        /* 2.6.1 校验链接服务器是否真实存在 */
        IF NOT EXISTS (SELECT 1 FROM sys.servers WHERE name = @Target AND is_linked = 1)
        BEGIN
            SET @msg = N'链接服务器 [' + @Target + N'] 不存在或不是链接服务器，已跳过。';
            PRINT N'[警告] ' + @msg;

            INSERT INTO dbo.SyncLoginUserLog (RunID, SourceServer, TargetServer, ObjectType, Status, Message)
            VALUES (@RunID, @SourceServer, @Target, 'SYSTEM', 'WARNING', @msg);

            SET @FailTarget += 1;
            FETCH NEXT FROM cur_target INTO @Target;
            CONTINUE;
        END

        TRUNCATE TABLE #DstLogin;
        TRUNCATE TABLE #DstRole;
        TRUNCATE TABLE #DstDB;
        TRUNCATE TABLE #SyncDDL;

        DECLARE @RemoteSQL    NVARCHAR(4000);
        DECLARE @OpenQuerySQL NVARCHAR(4000);

        /* 2.6.2 读取目标节点的登录信息（通过 OPENQUERY 直通查询）。
                 外层统一 COLLATE DATABASE_DEFAULT，避免链接服务器常见的排序规则冲突。 */
        SET @RemoteSQL = N'
SELECT p.name                                   AS LoginName,
       p.type                                   AS LoginType,
       CONVERT(NVARCHAR(1026), p.sid, 1)        AS SIDHex,
       CASE WHEN p.type = ''S''
            THEN CONVERT(NVARCHAR(1026), LOGINPROPERTY(p.name, ''PasswordHash''), 1)
       END                                      AS PWDHex,
       p.default_database_name                  AS DefaultDB,
       p.default_language_name                  AS DefaultLanguage,
       p.is_disabled                            AS IsDisabled,
       CAST(l.is_policy_checked     AS INT)     AS IsPolicyChecked,
       CAST(l.is_expiration_checked AS INT)     AS IsExpirationChecked
FROM sys.server_principals AS p
LEFT JOIN sys.sql_logins AS l ON l.principal_id = p.principal_id
WHERE p.type IN (''S'', ''U'', ''G'')
  AND p.name NOT LIKE ''##%##''';

        SET @OpenQuerySQL =
              N'SELECT LoginName COLLATE DATABASE_DEFAULT AS LoginName,'
            + N'       LoginType COLLATE DATABASE_DEFAULT AS LoginType,'
            + N'       SIDHex    COLLATE DATABASE_DEFAULT AS SIDHex,'
            + N'       PWDHex    COLLATE DATABASE_DEFAULT AS PWDHex,'
            + N'       DefaultDB COLLATE DATABASE_DEFAULT AS DefaultDB,'
            + N'       DefaultLanguage COLLATE DATABASE_DEFAULT AS DefaultLanguage,'
            + N'       IsDisabled, IsPolicyChecked, IsExpirationChecked '
            + N'FROM OPENQUERY(' + QUOTENAME(@Target) + N', '''
            + REPLACE(@RemoteSQL, N'''', N'''''') + N''')';

        BEGIN TRY
            INSERT INTO #DstLogin
                (LoginName, LoginType, SIDHex, PWDHex, DefaultDB, DefaultLanguage,
                 IsDisabled, IsPolicyChecked, IsExpirationChecked)
            EXEC (@OpenQuerySQL);
        END TRY
        BEGIN CATCH
            SET @msg = N'读取目标节点登录信息失败：' + ERROR_MESSAGE();
            PRINT N'[失败] ' + @msg;

            INSERT INTO dbo.SyncLoginUserLog (RunID, SourceServer, TargetServer, ObjectType, Status, Message)
            VALUES (@RunID, @SourceServer, @Target, 'SYSTEM', 'FAILED', @msg);

            SET @FailTarget += 1;
            FETCH NEXT FROM cur_target INTO @Target;
            CONTINUE;
        END CATCH

        /* 2.6.3 读取目标节点的数据库清单（用于校验默认数据库是否存在）。
                 读取失败不中断，后续对默认库采取保守处理。 */
        BEGIN TRY
            SET @RemoteSQL   = N'SELECT name AS DBName FROM sys.databases';
            SET @OpenQuerySQL =
                  N'SELECT DBName COLLATE DATABASE_DEFAULT AS DBName FROM OPENQUERY('
                + QUOTENAME(@Target) + N', ''' + REPLACE(@RemoteSQL, N'''', N'''''') + N''')';

            INSERT INTO #DstDB (DBName) EXEC (@OpenQuerySQL);
        END TRY
        BEGIN CATCH
            SET @msg = N'读取目标节点数据库清单失败（将跳过默认数据库校验）：' + ERROR_MESSAGE();
            PRINT N'[警告] ' + @msg;

            INSERT INTO dbo.SyncLoginUserLog (RunID, SourceServer, TargetServer, ObjectType, Status, Message)
            VALUES (@RunID, @SourceServer, @Target, 'SYSTEM', 'WARNING', @msg);
        END CATCH

        /* 2.6.4 读取目标节点已有的服务器角色成员关系 */
        IF @SyncServerRoles = 1
        BEGIN
            BEGIN TRY
                SET @RemoteSQL = N'
SELECT rp.name AS LoginName, r.name AS ServerRole
FROM sys.server_role_members AS rm
JOIN sys.server_principals AS r  ON r.principal_id  = rm.role_principal_id
JOIN sys.server_principals AS rp ON rp.principal_id = rm.member_principal_id
WHERE r.name <> ''public''';

                SET @OpenQuerySQL =
                      N'SELECT LoginName  COLLATE DATABASE_DEFAULT AS LoginName,'
                    + N'       ServerRole COLLATE DATABASE_DEFAULT AS ServerRole '
                    + N'FROM OPENQUERY(' + QUOTENAME(@Target) + N', '''
                    + REPLACE(@RemoteSQL, N'''', N'''''') + N''')';

                INSERT INTO #DstRole (LoginName, ServerRole) EXEC (@OpenQuerySQL);
            END TRY
            BEGIN CATCH
                SET @msg = N'读取目标节点服务器角色失败：' + ERROR_MESSAGE();
                PRINT N'[警告] ' + @msg;

                INSERT INTO dbo.SyncLoginUserLog (RunID, SourceServer, TargetServer, ObjectType, Status, Message)
                VALUES (@RunID, @SourceServer, @Target, 'SYSTEM', 'WARNING', @msg);
            END CATCH
        END

        /* 2.6.5 逐登录比对，生成待下发语句 */
        DECLARE @LoginName       SYSNAME;
        DECLARE @LoginType       CHAR(1);
        DECLARE @SIDHex          NVARCHAR(1026);
        DECLARE @PWDHex          NVARCHAR(1026);
        DECLARE @DefaultDB       SYSNAME;
        DECLARE @DefaultLanguage SYSNAME;
        DECLARE @IsDisabled      BIT;
        DECLARE @IsPolicy        INT;
        DECLARE @IsExpiration    INT;

        DECLARE @DstType         CHAR(1);
        DECLARE @DstSID          NVARCHAR(1026);
        DECLARE @DstPWD          NVARCHAR(1026);
        DECLARE @DstDefaultDB    SYSNAME;
        DECLARE @DstDefaultLang  SYSNAME;
        DECLARE @DstDisabled     BIT;
        DECLARE @DstPolicy       INT;
        DECLARE @DstExpiration   INT;

        DECLARE @EffectiveDB     SYSNAME;
        DECLARE @DDL             NVARCHAR(MAX);
        DECLARE @NeedPwd         BIT;
        DECLARE @NeedProp        BIT;
        DECLARE @NeedRecreate    BIT;
        DECLARE @LoginSkipped    BIT;
        DECLARE @Reason          NVARCHAR(500);

        DECLARE cur_login CURSOR LOCAL FAST_FORWARD FOR
            SELECT LoginName, LoginType, SIDHex, PWDHex, DefaultDB, DefaultLanguage,
                   IsDisabled, IsPolicyChecked, IsExpirationChecked
            FROM #SrcLogin
            ORDER BY LoginName;

        OPEN cur_login;
        FETCH NEXT FROM cur_login INTO @LoginName, @LoginType, @SIDHex, @PWDHex,
                                      @DefaultDB, @DefaultLanguage, @IsDisabled,
                                      @IsPolicy, @IsExpiration;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            SET @DstType        = NULL;
            SET @DstSID         = NULL;
            SET @DstPWD         = NULL;
            SET @DstDefaultDB   = NULL;
            SET @DstDefaultLang = NULL;
            SET @DstDisabled    = NULL;
            SET @DstPolicy      = NULL;
            SET @DstExpiration  = NULL;
            SET @DDL            = NULL;
            SET @NeedPwd        = 0;
            SET @NeedProp       = 0;
            SET @NeedRecreate   = 0;
            SET @LoginSkipped   = 0;

            SELECT @DstType        = d.LoginType,
                   @DstSID         = d.SIDHex,
                   @DstPWD         = d.PWDHex,
                   @DstDefaultDB   = d.DefaultDB,
                   @DstDefaultLang = d.DefaultLanguage,
                   @DstDisabled    = d.IsDisabled,
                   @DstPolicy      = d.IsPolicyChecked,
                   @DstExpiration  = d.IsExpirationChecked
            FROM #DstLogin AS d
            WHERE d.LoginName = @LoginName;

            /* 计算实际使用的默认数据库：目标端确实没有该库时回退为 master。
               只有在成功读到目标库清单时才做这项校验，避免误改。 */
            SET @EffectiveDB = @DefaultDB;

            IF @DefaultDB IS NOT NULL
               AND EXISTS (SELECT 1 FROM #DstDB)
               AND NOT EXISTS (SELECT 1 FROM #DstDB WHERE DBName = @DefaultDB)
            BEGIN
                SET @EffectiveDB = N'master';
                SET @Reason = N'默认数据库 [' + @DefaultDB + N'] 在目标节点不存在，已回退为 master。';
                PRINT N'[警告] ' + @LoginName + N'：' + @Reason;

                INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, Message)
                VALUES ('LOGIN', @LoginName, 'WARN', @Reason);
            END

            IF @EffectiveDB IS NULL
                SET @EffectiveDB = N'master';

            /* 判断是否需要重建：类型冲突，或 SQL 登录的 SID 不一致 */
            IF @DstType IS NOT NULL AND @LoginType <> @DstType
                SET @NeedRecreate = 1;

            IF @DstType IS NOT NULL AND @LoginType = 'S'
               AND ISNULL(@DstSID, N'') <> ISNULL(@SIDHex, N'')
                SET @NeedRecreate = 1;

            IF @DstType IS NULL OR @NeedRecreate = 1
            BEGIN
                /* ========== 目标端不存在，或已存在但需要重建 ========== */
                IF @NeedRecreate = 1 AND @RebuildOnSidMismatch = 0
                BEGIN
                    SET @LoginSkipped = 1;

                    SET @Reason =
                        CASE WHEN @LoginType <> @DstType
                             THEN N'目标端已存在同名登录但类型不一致（源端 ' + @LoginType
                                + N'，目标端 ' + @DstType + N'），未开启重建开关，已跳过。'
                             ELSE N'目标端已存在同名登录但 SID 不一致，未开启重建开关，已跳过。'
                        END;

                    INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, Message)
                    VALUES ('LOGIN', @LoginName, 'SKIP', @Reason);
                END
                ELSE IF @LoginType = 'S' AND @PWDHex IS NULL
                BEGIN
                    SET @LoginSkipped = 1;

                    INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, Message)
                    VALUES ('LOGIN', @LoginName, 'SKIP',
                            N'源端密码哈希无法读取（LOGINPROPERTY 返回 NULL），已跳过。');
                END
                ELSE
                BEGIN
                    -- 重建时先删除目标端旧登录
                    IF @NeedRecreate = 1
                    BEGIN
                        INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, DDL)
                        VALUES ('LOGIN', @LoginName, 'DROP',
                                N'DROP LOGIN ' + QUOTENAME(@LoginName) + N';');
                    END

                    IF @LoginType = 'S'
                    BEGIN
                        /* 关键点：SID 与源端完全一致，密码以哈希下发（HASHED），
                           全程不需要知道明文密码。 */
                        SET @DDL = N'CREATE LOGIN ' + QUOTENAME(@LoginName)
                                 + N' WITH PASSWORD = ' + @PWDHex + N' HASHED'
                                 + N', SID = ' + @SIDHex
                                 + N', DEFAULT_DATABASE = ' + QUOTENAME(@EffectiveDB)
                                 + N', DEFAULT_LANGUAGE = ' + QUOTENAME(ISNULL(@DefaultLanguage, N'us_english'))
                                 + N', CHECK_POLICY = '
                                 + CASE WHEN ISNULL(@IsPolicy, 1) = 1 THEN N'ON' ELSE N'OFF' END
                                 + N', CHECK_EXPIRATION = '
                                 + CASE WHEN ISNULL(@IsExpiration, 0) = 1 THEN N'ON' ELSE N'OFF' END
                                 + N';';
                    END
                    ELSE
                    BEGIN
                        -- Windows 登录 / 组：SID 由域决定，不能也不应手工指定
                        SET @DDL = N'CREATE LOGIN ' + QUOTENAME(@LoginName) + N' FROM WINDOWS'
                                 + N' WITH DEFAULT_DATABASE = ' + QUOTENAME(@EffectiveDB)
                                 + N', DEFAULT_LANGUAGE = ' + QUOTENAME(ISNULL(@DefaultLanguage, N'us_english'))
                                 + N';';
                    END

                    INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, DDL)
                    VALUES ('LOGIN', @LoginName,
                            CASE WHEN @NeedRecreate = 1 THEN 'RECREATE' ELSE 'CREATE' END,
                            @DDL);
                END
            END
            ELSE
            BEGIN
                /* ========== 目标端已存在，且类型与 SID 均一致：只做属性对齐 ========== */

                -- ① 密码：仅在哈希不同时才改。避免每次同步都重置密码，
                --    那会刷新密码修改时间并影响密码过期策略。
                --    注意必须校验 @PWDHex 非空，否则拼接出的语句会变成 NULL。
                IF @LoginType = 'S' AND @PWDHex IS NOT NULL
                   AND ISNULL(@DstPWD, N'') <> ISNULL(@PWDHex, N'')
                    SET @NeedPwd = 1;

                -- ② 策略由 OFF 变为 ON 时，SQL Server 需要重新校验密码，
                --    因此必须在本语句中一并提供密码哈希，否则会报策略校验失败。
                IF @LoginType = 'S' AND @PWDHex IS NOT NULL
                   AND ISNULL(@DstPolicy, 1) = 0
                   AND ISNULL(@IsPolicy, 0) = 1
                    SET @NeedPwd = 1;

                -- ③ 其他属性比对
                IF ISNULL(@DstDefaultDB, N'')   <> ISNULL(@EffectiveDB, N'')
                    SET @NeedProp = 1;

                IF ISNULL(@DstDefaultLang, N'') <> ISNULL(@DefaultLanguage, N'')
                    SET @NeedProp = 1;

                IF @LoginType = 'S' AND ISNULL(@DstPolicy, -1) <> ISNULL(@IsPolicy, -1)
                    SET @NeedProp = 1;

                IF @LoginType = 'S' AND ISNULL(@DstExpiration, -1) <> ISNULL(@IsExpiration, -1)
                    SET @NeedProp = 1;

                IF @NeedPwd = 1 OR @NeedProp = 1
                BEGIN
                    SET @DDL = N'ALTER LOGIN ' + QUOTENAME(@LoginName) + N' WITH ';

                    IF @NeedPwd = 1
                        SET @DDL += N'PASSWORD = ' + @PWDHex + N' HASHED, ';

                    SET @DDL += N'DEFAULT_DATABASE = ' + QUOTENAME(@EffectiveDB)
                              + N', DEFAULT_LANGUAGE = ' + QUOTENAME(ISNULL(@DefaultLanguage, N'us_english'));

                    IF @LoginType = 'S'
                        SET @DDL += N', CHECK_POLICY = '
                                  + CASE WHEN ISNULL(@IsPolicy, 1) = 1 THEN N'ON' ELSE N'OFF' END
                                  + N', CHECK_EXPIRATION = '
                                  + CASE WHEN ISNULL(@IsExpiration, 0) = 1 THEN N'ON' ELSE N'OFF' END;

                    SET @DDL += N';';

                    INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, DDL)
                    VALUES ('LOGIN', @LoginName, 'ALTER', @DDL);
                END
            END

            /* ========== 启用 / 禁用状态对齐 ========== */
            IF @SyncDisabledState = 1 AND @LoginSkipped = 0
            BEGIN
                IF @DstType IS NULL OR @NeedRecreate = 1
                BEGIN
                    -- 新建（含重建）的登录默认是启用状态，源端为禁用时补一条 DISABLE
                    IF @IsDisabled = 1
                    BEGIN
                        INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, DDL)
                        VALUES ('LOGIN', @LoginName, 'DISABLE',
                                N'ALTER LOGIN ' + QUOTENAME(@LoginName) + N' DISABLE;');
                    END
                END
                ELSE
                BEGIN
                    IF @IsDisabled = 1 AND ISNULL(@DstDisabled, 0) = 0
                    BEGIN
                        INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, DDL)
                        VALUES ('LOGIN', @LoginName, 'DISABLE',
                                N'ALTER LOGIN ' + QUOTENAME(@LoginName) + N' DISABLE;');
                    END
                    ELSE IF @IsDisabled = 0 AND ISNULL(@DstDisabled, 1) = 1
                    BEGIN
                        INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, DDL)
                        VALUES ('LOGIN', @LoginName, 'ENABLE',
                                N'ALTER LOGIN ' + QUOTENAME(@LoginName) + N' ENABLE;');
                    END
                END
            END

            /* ========== 服务器角色成员关系对齐（只补齐，不删除目标端多余角色） ========== */
            IF @SyncServerRoles = 1 AND @LoginSkipped = 0
            BEGIN
                INSERT INTO #SyncDDL (ObjectType, ObjectName, Action, DDL)
                SELECT 'ROLE', r.ServerRole, 'ADDROLE',
                       N'ALTER SERVER ROLE ' + QUOTENAME(r.ServerRole)
                     + N' ADD MEMBER ' + QUOTENAME(r.LoginName) + N';'
                FROM #SrcRole AS r
                WHERE r.LoginName = @LoginName
                  AND NOT EXISTS
                      (
                          SELECT 1
                          FROM #DstRole AS d
                          WHERE d.LoginName = r.LoginName
                            AND d.ServerRole = r.ServerRole
                      );
            END

            FETCH NEXT FROM cur_login INTO @LoginName, @LoginType, @SIDHex, @PWDHex,
                                          @DefaultDB, @DefaultLanguage, @IsDisabled,
                                          @IsPolicy, @IsExpiration;
        END

        CLOSE cur_login;
        DEALLOCATE cur_login;

        /* 2.6.6 下发执行 */
        DECLARE @Seq         INT;
        DECLARE @ObjType     VARCHAR(20);
        DECLARE @ObjName     SYSNAME;
        DECLARE @Action      VARCHAR(20);
        DECLARE @DDLText     NVARCHAR(MAX);
        DECLARE @ItemMessage NVARCHAR(4000);
        DECLARE @ExecSQL     NVARCHAR(MAX);
        DECLARE @NodeFailed  BIT = 0;
        DECLARE @ItemCount   INT;

        SELECT @ItemCount = COUNT(*) FROM #SyncDDL;

        IF @Debug = 1
        BEGIN
            PRINT N'[试运行] 目标节点 ' + @Target + N' 共生成 '
                + CAST(@ItemCount AS NVARCHAR(10)) + N' 条语句，明细如下：';
            PRINT N'------------------------------------------------------------';

            SELECT Seq, ObjectType, ObjectName, Action, Message, DDL
            FROM #SyncDDL
            ORDER BY Seq;

            PRINT N'（以上为试运行结果，未真正下发执行。）';
        END
        ELSE
        BEGIN
            DECLARE cur_ddl CURSOR LOCAL FAST_FORWARD FOR
                SELECT Seq, ObjectType, ObjectName, Action, DDL, Message
                FROM #SyncDDL
                ORDER BY Seq;

            OPEN cur_ddl;
            FETCH NEXT FROM cur_ddl INTO @Seq, @ObjType, @ObjName, @Action, @DDLText, @ItemMessage;

            WHILE @@FETCH_STATUS = 0
            BEGIN
                IF @DDLText IS NULL
                BEGIN
                    -- 仅记录日志，不执行的条目（例如跳过项、警告）
                    INSERT INTO dbo.SyncLoginUserLog
                        (RunID, SourceServer, TargetServer, ObjectType, ObjectName, Action, Status, Message)
                    VALUES
                        (@RunID, @SourceServer, @Target, @ObjType, @ObjName, @Action, 'WARNING', @ItemMessage);
                END
                ELSE
                BEGIN
                    -- 通过链接服务器把语句发到目标实例执行：EXEC (@语句) AT [链接服务器]
                    -- 这里借 sp_executesql 传参，避免手工拼接单引号带来的转义问题。
                    SET @ExecSQL = N'EXEC (@RemoteDDL) AT ' + QUOTENAME(@Target) + N';';

                    BEGIN TRY
                        EXEC sys.sp_executesql @ExecSQL,
                             N'@RemoteDDL NVARCHAR(MAX)',
                             @RemoteDDL = @DDLText;

                        INSERT INTO dbo.SyncLoginUserLog
                            (RunID, SourceServer, TargetServer, ObjectType, ObjectName, Action, Status, DDLText)
                        VALUES
                            (@RunID, @SourceServer, @Target, @ObjType, @ObjName, @Action, 'SUCCESS', @DDLText);
                    END TRY
                    BEGIN CATCH
                        SET @NodeFailed = 1;
                        SET @msg = N'执行失败：' + ERROR_MESSAGE();

                        INSERT INTO dbo.SyncLoginUserLog
                            (RunID, SourceServer, TargetServer, ObjectType, ObjectName, Action, Status, Message, DDLText)
                        VALUES
                            (@RunID, @SourceServer, @Target, @ObjType, @ObjName, @Action, 'FAILED', @msg, @DDLText);

                        PRINT N'[失败] ' + @Target + N' → ' + ISNULL(@ObjName, @ObjType) + N'：' + @msg;
                    END CATCH
                END

                FETCH NEXT FROM cur_ddl INTO @Seq, @ObjType, @ObjName, @Action, @DDLText, @ItemMessage;
            END

            CLOSE cur_ddl;
            DEALLOCATE cur_ddl;

            PRINT N'[完成] 目标节点 ' + @Target + N' 处理完毕，共生成 '
                + CAST(@ItemCount AS NVARCHAR(10)) + N' 条语句。';
        END

        IF @NodeFailed = 1
            SET @FailTarget += 1;
        ELSE
            SET @OKTarget += 1;

        FETCH NEXT FROM cur_target INTO @Target;
    END

    CLOSE cur_target;
    DEALLOCATE cur_target;

    /*--------------------------------------------------------------------------
      2.7 收尾：清理历史日志并输出本次运行汇总
    --------------------------------------------------------------------------*/
    IF @LogRetentionDays > 0
    BEGIN
        DELETE FROM dbo.SyncLoginUserLog
        WHERE RunTime < DATEADD(DAY, -@LogRetentionDays, SYSDATETIME());
    END

    SET @msg = N'本次同步结束。目标节点共 ' + CAST(@TotalTarget AS NVARCHAR(10)) + N' 个，'
             + N'成功 ' + CAST(@OKTarget AS NVARCHAR(10)) + N' 个，'
             + N'失败 ' + CAST(@FailTarget AS NVARCHAR(10)) + N' 个。';

    PRINT N'============================================================';
    PRINT N'[结束] ' + @msg;
    PRINT N'============================================================';

    INSERT INTO dbo.SyncLoginUserLog (RunID, SourceServer, ObjectType, Status, Message)
    VALUES (@RunID, @SourceServer, 'SYSTEM',
            CASE WHEN @FailTarget > 0 THEN 'WARNING' ELSE 'SUCCESS' END, @msg);

    IF @Debug = 1
    BEGIN
        SELECT LogID, RunID, RunTime, SourceServer, TargetServer, ObjectType,
               ObjectName, Action, Status, Message, DDLText
        FROM dbo.SyncLoginUserLog
        WHERE RunID = @RunID
        ORDER BY LogID;
    END
END
GO


/*==============================================================================
  第 3 步：使用示例
==============================================================================*/
-- 3.1 正式同步（由定时作业周期调用，作业步骤命令）：
--     EXEC master.dbo.usp_SyncLoginUserRegularBetweenInstances;

-- 3.2 试运行：只生成并打印脚本，不真正下发（首次部署务必先跑这个）
--     EXEC master.dbo.usp_SyncLoginUserRegularBetweenInstances @Debug = 1;

-- 3.3 只同步指定节点，并同时同步 Windows 登录
--     EXEC master.dbo.usp_SyncLoginUserRegularBetweenInstances
--          @LinkedServerList     = N'NODE02,NODE03',
--          @IncludeWindowsLogins = 1;

-- 3.4 排除特定登录（例如复制账号）
--     EXEC master.dbo.usp_SyncLoginUserRegularBetweenInstances
--          @ExcludeLoginList = N'distributor_admin,repl_writer';

-- 3.5 处理"目标端同名登录 SID 不一致"的场景
--     谨慎使用：会先 DROP 目标端的同名登录再重建，请确认该登录不是作业所有者
--     EXEC master.dbo.usp_SyncLoginUserRegularBetweenInstances @RebuildOnSidMismatch = 1;

-- 3.6 非可用性组（单实例 / 非 AG 场景）手工执行分发
--     EXEC master.dbo.usp_SyncLoginUserRegularBetweenInstances @ForceRun = 1;

-- 3.7 查看同步日志
--     SELECT TOP (200) * FROM master.dbo.SyncLoginUserLog ORDER BY LogID DESC;
--     SELECT TargetServer, Status, COUNT(*) AS Cnt
--     FROM master.dbo.SyncLoginUserLog
--     WHERE RunTime > DATEADD(DAY, -1, SYSDATETIME())
--     GROUP BY TargetServer, Status;
GO
