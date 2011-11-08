<!--
<DOCUMENTATION>
	<MODULE_NAME>adoJavas.inc
	<AUTHOR>a-ricksa
	<DESCRIPTION>
	The <b>adoJavas.inc</b> module is used to define a set of ADO constants for JavaScript routines.  No database access is used in this module.
</DOCUMENTATION>

<TECHNOTE NAME="(No code highlighted)">
' This module does not have any highlighted code
</TECHNOTE>
-->

<%
//--------------------------------------------------------------------
// Microsoft ADO
//
// (c) 1996 Microsoft Corporation.  All Rights Reserved.
//
//
//
// ADO constants include file for JavaScript
//
//--------------------------------------------------------------------

//---- CursorTypeEnum Values ----
var adOpenForwardOnly = 0;
var adOpenKeyset = 1;
var adOpenDynamic = 2;
var adOpenStatic = 3;

//---- CursorOptionEnum Values ----
var adHoldRecords = 0x00000100;
var adMovePrevious = 0x00000200;
var adAddNew = 0x01000400;
var adDelete = 0x01000800;
var adUpdate = 0x01008000;
var adBookmark = 0x00002000;
var adApproxPosition = 0x00004000;
var adUpdateBatch = 0x00010000;
var adResync = 0x00020000;

//---- LockTypeEnum Values ----
var adLockReadOnly = 1;
var adLockPessimistic = 2;
var adLockOptimistic = 3;
var adLockBatchOptimistic = 4;

//---- CursorLocationEnum Values ----
var adUseClient = 1;
var adUseServer = 2;
var adUseClientBatch = 3;

//---- DataTypeEnum Values ----
var adEmpty = 0;
var adTinyInt = 16;
var adSmallInt = 2;
var adInteger = 3;
var adBigInt = 20;
var adUnsignedTinyInt = 17;
var adUnsignedSmallInt = 18;
var adUnsignedInt = 19;
var adUnsignedBigInt = 21;
var adSingle = 4;
var adDouble = 5;
var adCurrency = 6;
var adDecimal = 14;
var adNumeric = 131;
var adBoolean = 11;
var adError = 10;
var adUserDefined = 132;
var adVariant = 12;
var adIDispatch = 9;
var adIUnknown = 13;
var adGUID = 72;
var adDate = 7;
var adDBDate = 133;
var adDBTime = 134;
var adDBTimeStamp = 135;
var adBSTR = 8;
var adChar = 129;
var adVarChar = 200;
var adLongVarChar = 201;
var adWChar = 130;
var adVarWChar = 202;
var adLongVarWChar = 203;
var adBinary = 128;
var adVarBinary = 204;
var adLongVarBinary = 205;

//---- ConnectPromptEnum Values ----
var adPromptAlways = 1;
var adPromptComplete = 2;
var adPromptCompleteRequired = 3;
var adPromptNever = 4;

//---- ConnectModeEnum Values ----
var adModeUnknown = 0;
var adModeRead = 1;
var adModeWrite = 2;
var adModeReadWrite = 3;
var adModeShareDenyRead = 4;
var adModeShareDenyWrite = 8;
var adModeShareExclusive = 0xc;
var adModeShareDenyNone = 0x10;

//---- IsolationLevelEnum Values ----
var adXactUnspecified = 0xffffffff;
var adXactChaos = 0x00000010;
var adXactReadUncommitted = 0x00000100;
var adXactBrowse = 0x00000100;
var adXactCursorStability = 0x00001000;
var adXactReadCommitted = 0x00001000;
var adXactRepeatableRead = 0x00010000;
var adXactSerializable = 0x00100000;
var adXactIsolated = 0x00100000;

//---- XactAttributeEnum Values ----
var adXactPollAsync = 2;
var adXactPollSyncPhaseOne = 4;
var adXactCommitRetaining = 0x00020000;
var adXactAbortRetaining = 0x00040000;
var adXactAbortAsync = 0x00080000;

//---- FieldAttributeEnum Values ----
var adFldBookmark = 0x00000001;
var adFldMayDefer = 0x00000002;
var adFldUpdatable = 0x00000004;
var adFldUnknownUpdatable = 0x00000008;
var adFldFixed = 0x00000010;
var adFldIsNullable = 0x00000020;
var adFldMayBeNull = 0x00000040;
var adFldLong = 0x00000080;
var adFldRowID = 0x00000100;
var adFldRowVersion = 0x00000200;
var adFldCacheDeferred = 0x00001000;

//---- EditModeEnum Values ----
var adEditNone = 0x0000;
var adEditInProgress = 0x0001;
var adEditAdd = 0x0002;

//---- RecordStatusEnum Values ----
var adRecOK = 0x0000000;
var adRecNew = 0x0000001;
var adRecModified = 0x0000002;
var adRecDeleted = 0x0000004;
var adRecUnmodified = 0x0000008;
var adRecInvalid = 0x0000010;
var adRecMultipleChanges = 0x0000040;
var adRecPendingChanges = 0x0000080;
var adRecCanceled = 0x0000100;
var adRecCantRelease = 0x0000400;
var adRecConcurrencyViolation = 0x0000800;
var adRecIntegrityViolation = 0x0001000;
var adRecMaxChangesExceeded = 0x0002000;
var adRecObjectOpen = 0x0004000;
var adRecOutOfMemory = 0x0008000;
var adRecPermissionDenied = 0x0010000;
var adRecSchemaViolation = 0x0020000;
var adRecDBDeleted = 0x0040000;

//---- GetRowsOptionEnum Values ----
var adGetRowsRest = -1;

//---- PositionEnum Values ----
var adPosUnknown = -1;
var adPosBOF = -2;
var adPosEOF = -3;

//---- AffectEnum Values ----
var adAffectCurrent = 1;
var adAffectGroup = 2;
var adAffectAll = 3;

//---- FilterGroupEnum Values ----
var adFilterNone = 0;
var adFilterPendingRecords = 1;
var adFilterAffectedRecords = 2;
var adFilterFetchedRecords = 3;

//---- PropertyAttributesEnum Values ----
var adPropNotSupported = 0x0000;
var adPropRequired = 0x0001;
var adPropOptional = 0x0002;
var adPropRead = 0x0200;
var adPropWrite = 0x0400;

//---- ErrorValueEnum Values ----
var adErrInvalidArgument = 0xbb9;
var adErrNoCurrentRecord = 0xbcd;
var adErrIllegalOperation = 0xc93;
var adErrInTransaction = 0xcae;
var adErrFeatureNotAvailable = 0xcb3;
var adErrItemNotFound = 0xcc1;
var adErrObjectNotSet = 0xd5c;
var adErrDataConversion = 0xd5d;
var adErrObjectClosed = 0xe78;
var adErrObjectOpen = 0xe79;
var adErrProviderNotFound = 0xe7a;
var adErrBoundToCommand = 0xe7b;

//---- ParameterAttributesEnum Values ----
var adParamSigned = 0x0010;
var adParamNullable = 0x0040;
var adParamLong = 0x0080;

//---- ParameterDirectionEnum Values ----
var adParamUnknown = 0x0000;
var adParamInput = 0x0001;
var adParamOutput = 0x0002;
var adParamInputOutput = 0x0003;
var adParamReturnValue = 0x0004;

//---- CommandTypeEnum Values ----
var adCmdUnknown = 0;
var adCmdText = 0x0001;
var adCmdTable = 0x0002;
var adCmdStoredProc = 0x0004;
%>
