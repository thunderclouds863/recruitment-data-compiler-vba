Attribute VB_Name = "m01_PublicDeclarations"
Option Explicit

' ============================================================
' PUBLIC DECLARATIONS
' ============================================================

Public Const SCHEMA_INTEGRATION_VERSION As String = _
    "Schema Adjustment Batch 03"

Public Const SCHEMA_DB_FPTK As String = _
    "FPTK"

Public Const SCHEMA_DB_KODE_POSISI As String = _
    "DB Kode Posisi"

Public Const SCHEMA_DB_SOURCING As String = _
    "DB Sourcing"

' Progress compiler.
Public gTotalSteps As Long
Public gCurrentStep As Long
Public gShowProgress As Boolean

' External/source Excel application.
Public gSourceApp As Object

' Schema-aware compile state.
'
' Variabel ini dapat dibaca oleh form UPDATE ADJUSTMENT
' agar adjustment tidak dijalankan ketika compiler aktif.
Public gSchemaCompileInProgress As Boolean
Public gSchemaCurrentStage As String
Public gSchemaCurrentFolder As String
Public gSchemaLastError As String


