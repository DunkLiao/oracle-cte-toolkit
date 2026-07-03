try:
    import oracledb
except ModuleNotFoundError:
    class _MissingOracleDb:
        def connect(self, *args, **kwargs):
            raise ModuleNotFoundError(
                "oracledb is required for Oracle connections. "
                "Install it with: pip install oracledb"
            )

    oracledb = _MissingOracleDb()


def connect(username, password, dsn):
    return oracledb.connect(
        user=username,
        password=password,
        dsn=dsn,
    )


def wrap_metadata_sql(sql_text):
    sql_text = sql_text.strip().rstrip(";")

    return f"""
        SELECT *
        FROM (
            {sql_text}
        ) X
        WHERE 1=0
        """


def oracle_type_from_description(col):
    dtype = str(col[1]).upper()
    length = col[3]
    precision = col[4]
    scale = col[5]

    if "VARCHAR" in dtype:
        return f"VARCHAR2({length})"

    if "CHAR" in dtype:
        return f"CHAR({length})"

    if "NUMBER" in dtype:
        if precision:
            if scale and scale > 0:
                return f"NUMBER({precision},{scale})"
            return f"NUMBER({precision})"
        return "NUMBER"

    if "DATE" in dtype:
        return "DATE"

    if "TIMESTAMP" in dtype:
        return "TIMESTAMP"

    if "CLOB" in dtype:
        return "CLOB"

    if "BLOB" in dtype:
        return "BLOB"

    return dtype
