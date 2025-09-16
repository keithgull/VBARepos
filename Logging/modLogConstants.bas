Attribute VB_Name = "modLogConstants"
' Module: LoggerConstants
Option Explicit

' ���O���x���萔�i�\���p�Ɍ������ς݁j
Public Const LOGLEVEL_INFO  As String = "INFO "
Public Const LOGLEVEL_DEBUG As String = "DEBUG"
Public Const LOGLEVEL_WARN  As String = "WARN "
Public Const LOGLEVEL_ERROR As String = "ERROR"
Public Const LOGLEVEL_TRACE As String = "TRACE"

' ���O���x���L�[�i��������p�j
'Public Const LOGKEY_INFO  As String = "INFO"
'Public Const LOGKEY_DEBUG As String = "DEBUG"
'Public Const LOGKEY_WARN  As String = "WARN"
'Public Const LOGKEY_ERROR As String = "ERROR"
'Public Const LOGKEY_TRACE As String = "TRACE"

' Enum��`�F���O���
Public Enum LOGGER_TYPE
    LOGGER_TYPE_DEBUGPRINT = 0
    LOGGER_TYPE_LOGFILE = 1
    LOGGER_TYPE_LOGSHEET = 2
End Enum
