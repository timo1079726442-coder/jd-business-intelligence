"""统一流水线状态与稳定退出码定义。"""

# 成功状态。SUCCESS_EMPTY 表示平台真实返回空表，不是失败。
SUCCESS_WITH_DATA = "SUCCESS_WITH_DATA"
SUCCESS_EMPTY = "SUCCESS_EMPTY"
# dry-run 已完成计划和本地鉴权检查，但未请求平台或写业务库。
DRY_RUN_READY = "DRY_RUN_READY"

# 失败状态。名称会写入 etl_task_run 和 JSON Summary，供影刀稳定判断。
FAILED_AUTH = "FAILED_AUTH"
FAILED_REQUEST = "FAILED_REQUEST"
FAILED_PARSE = "FAILED_PARSE"
FAILED_DATABASE = "FAILED_DATABASE"
FAILED_VALIDATION = "FAILED_VALIDATION"

SUCCESS_STATUSES = {SUCCESS_WITH_DATA, SUCCESS_EMPTY, DRY_RUN_READY}
FAILURE_STATUSES = {
    FAILED_AUTH,
    FAILED_REQUEST,
    FAILED_PARSE,
    FAILED_DATABASE,
    FAILED_VALIDATION,
}

# 新入口专用退出码。保留旧脚本 0/1/2/3 的兼容性，但影刀应只判断这些固定值。
EXIT_SUCCESS = 0
EXIT_PARTIAL_FAILURE = 10
EXIT_AUTH_FAILURE = 20
EXIT_REQUEST_FAILURE = 30
EXIT_PARSE_FAILURE = 40
EXIT_DATABASE_FAILURE = 50
EXIT_VALIDATION_FAILURE = 60
EXIT_PARAMETER_ERROR = 70


def choose_exit_code(task_statuses: list[str]) -> int:
    """按失败类型优先级选择整个 Run 的稳定退出码。"""
    statuses = set(task_statuses)
    if not statuses or statuses.issubset(SUCCESS_STATUSES):
        return EXIT_SUCCESS
    if FAILED_AUTH in statuses:
        return EXIT_AUTH_FAILURE
    if FAILED_DATABASE in statuses:
        return EXIT_DATABASE_FAILURE
    if FAILED_VALIDATION in statuses:
        return EXIT_VALIDATION_FAILURE
    if FAILED_PARSE in statuses:
        return EXIT_PARSE_FAILURE
    if FAILED_REQUEST in statuses:
        return EXIT_REQUEST_FAILURE
    return EXIT_PARTIAL_FAILURE
