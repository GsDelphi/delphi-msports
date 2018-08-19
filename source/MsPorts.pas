{******************************************************************************}

{ XXXX API interface Unit for Object Pascal                                    }

 { Portions created by Microsoft are Copyright (C) 1995-2008 Microsoft          }
 { Corporation. All Rights Reserved.                                            }

 { Portions created by XXXXXXXXXXXXXXXXX are Copyright (C) xxxx-xxxx            }
 { XXXXXXXXXXXXXXXXX. All Rights Reserved.                                      }

{ Obtained through: Joint Endeavour of Delphi Innovators (Project JEDI)        }

 { You may retrieve the latest version of this file at the Project JEDI         }
 { APILIB home page, located at http://jedi-apilib.sourceforge.net              }

 { The contents of this file are used with permission, subject to the Mozilla   }
 { Public License Version 1.1 (the "License"); you may not use this file except }
 { in compliance with the License. You may obtain a copy of the License at      }
 { http://www.mozilla.org/MPL/MPL-1.1.html                                      }

 { Software distributed under the License is distributed on an "AS IS" basis,   }
 { WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for }
 { the specific language governing rights and limitations under the License.    }

 { Alternatively, the contents of this file may be used under the terms of the  }
 { GNU Lesser General Public License (the  "LGPL License"), in which case the   }
 { provisions of the LGPL License are applicable instead of those above.        }
 { If you wish to allow use of your version of this file only under the terms   }
 { of the LGPL License and not to allow others to use your version of this file }
 { under the MPL, indicate your decision by deleting  the provisions above and  }
 { replace  them with the notice and other provisions required by the LGPL      }
 { License.  If you do not delete the provisions above, a recipient may use     }
 { your version of this file under either the MPL or the LGPL License.          }

{ For more information about the LGPL: http://www.gnu.org/copyleft/lesser.html }

{******************************************************************************}

unit MsPorts;

{.$DEFINE SERIAL_STATIC_LINK}
{.$DEFINE SERIAL_USE_DELPHI_TYPES}
{.$DEFINE SERIAL_USE_JWA_SINLGE_UNITS}
{.$DEFINE SERIAL_ADVANCED_SETTINGS}

{$HPPEMIT ''}
{$HPPEMIT '#include "msports.h"'}
{$HPPEMIT ''}

{.$I ..\Includes\JediAPILib.inc}

interface

{$WEAKPACKAGEUNIT ON}

{$IFNDEF SERIAL_USE_DELPHI_TYPES}

  {$IFNDEF SERIAL_USE_JWA_SINLGE_UNITS}

uses
  JwaWindows;

  {$ELSE SERIAL_USE_JWA_SINLGE_UNITS}

uses
  JwaWinBase,
  JwaWinType;

  {$ENDIF SERIAL_USE_JWA_SINLGE_UNITS}

{$ELSE SERIAL_USE_DELPHI_TYPES}

uses
  Windows,
  SysUtils;

type
  EMsPortsError = class(Exception);
  ELoadLibraryError = class(EMsPortsError);
  EGetProcAddressError = class(EMsPortsError);

{$ENDIF SERIAL_USE_DELPHI_TYPES}


type
  LONG = Longint;


{$IFDEF SERIAL_ADVANCED_SETTINGS}

  {$IFDEF SERIAL_USE_JWA_SINLGE_UNITS}
    {$DEFINE SERIAL_SETUP_API_TYPES}
  {$ENDIF SERIAL_USE_JWA_SINLGE_UNITS}

  {$IFDEF SERIAL_USE_DELPHI_TYPES}
    {$DEFINE SERIAL_SETUP_API_TYPES}
  {$ENDIF SERIAL_USE_DELPHI_TYPES}

  
  {$IFDEF SERIAL_USE_DELPHI_TYPES}

  SizeUInt = {$IFDEF COMPILER16_UP} NativeUInt {$ELSE} Longword {$ENDIF};

  ULONG_PTR = {$IFDEF COMPILER12_UP} Windows.ULONG_PTR {$ELSE} SizeUInt {$ENDIF};
  {$EXTERNALSYM ULONG_PTR}

  {$ENDIF SERIAL_USE_DELPHI_TYPES}


  {$IFDEF SERIAL_SETUP_API_TYPES}

  // Define type for reference to device information set
  HDEVINFO = Pointer;
  {$EXTERNALSYM HDEVINFO}


  // Device information structure (references a device instance
  // that is a member of a device information set)
  PSPDevInfoData = ^TSPDevInfoData;

  SP_DEVINFO_DATA = packed record
    cbSize:    DWORD;
    ClassGuid: TGUID;
    DevInst:   DWORD; // DEVINST handle
    Reserved:  ULONG_PTR;
  end;
  {$EXTERNALSYM SP_DEVINFO_DATA}

  TSPDevInfoData = SP_DEVINFO_DATA;

  {$ENDIF SERIAL_SETUP_API_TYPES}

(*++

Routine Description:

    Displays the advanced properties dialog for the COM port specified by
    DeviceInfoSet and DeviceInfoData.

Arguments:

    ParentHwnd  - the parent window of the window to be displayed

    DeviceInfoSet, DeviceInfoData - SetupDi structures representing the COM port

Return Value:

    ERROR_SUCCESS if the dialog was shown

  --*)
function SerialDisplayAdvancedSettings(const ParentHwnd: HWND;
  const DeviceInfoSet: HDEVINFO;
  const DeviceInfoData: PSPDevInfoData): LONG cdecl stdcall;

(*++

Routine Description:

    Prototype to allow serial port vendors to override the advanced dialog
    represented by the COM port specified by DeviceInfoSet and DeviceInfoData.

    To override the advanced page, place a value named EnumAdvancedDialog under
    the same key in which you would put your EnumPropPages32 value.  The format
    of the value is exactly the same as Enum...32 as well.

Arguments:

    ParentHwnd  - the parent window of the window to be displayed

    HidePollingUI - If TRUE, hide all UI that deals with polling.

    DeviceInfoSet, DeviceInfoData - SetupDi structures representing the COM port

    Reserved - Unused

Return Value:

    TRUE if the user pressed OK, FALSE if Cancel was pressed
--*)

type
  PPORT_ADVANCED_DIALOG = function(const ParentHwnd: HWND;
    const HidePollingUI: BOOL; const DeviceInfoSet: HDEVINFO;
    const DeviceInfoData: PSPDevInfoData; const Reserved: Pointer): BOOL;

{$endif}

type
  HCOMDB  = LONG;
  PHCOMDB = ^HCOMDB;

const
  HCOMDB_INVALID_HANDLE_VALUE = HCOMDB(INVALID_HANDLE_VALUE);


// Minimum through maximum number of COM names arbitered

const
  COMDB_MIN_PORTS_ARBITRATED = 256;
  COMDB_MAX_PORTS_ARBITRATED = 4096;

(*++

Routine Description:

    Opens name data base, and returns a handle to be used in future calls.

Arguments:

    None.

Return Value:

    INVALID_HANDLE_VALUE if the call fails, otherwise a valid handle

    If INVALID_HANDLE_VALUE, call GetLastError() to get details (??)

--*)
function ComDBOpen(var aPHComDB: HCOMDB): LONG stdcall;


(*++

Routine Description:

    frees a handle to the database returned from OpenComPortDataBase

Arguments:

    Handle returned from OpenComPortDataBase.

Return Value:

    None

--*)
function ComDBClose(aHComDB: HCOMDB): LONG stdcall;


const
  CDB_REPORT_BITS  = $0;
  CDB_REPORT_BYTES = $1;

(*++

Routine Description:

    if Buffer is NULL, than MaxPortsReported will contain the max number of ports
        the DB will report (this value is NOT the number of bytes need for Buffer).
        ReportType is ignored in this case.

    if ReportType == CDB_REPORT_BITS
        returns a bit array indicating if a comX name is claimed.
        ie, Bit 0 of Byte 0 is com1, bit 1 of byte 0 is com2 and so on.

        BufferSize >= MaxPortsReported / 8


    if ReportType == CDB_REPORT_BYTES
        returns a byte array indicating if a comX name is claimed.  Zero unused, non zero
        used, ie, byte 0 is com1, byte 1 is com2, etc

        BufferSize >= MaxPortsReported

Arguments:

    Handle returned from OpenComPortDataBase.

    Buffer pointes to memory to place bit array

    BufferSize   Size of buffer in bytes

    MaxPortsReported    Pointer to DWORD that holds the number of bytes in buffer filled in

Return Value:

    returns ERROR_SUCCESS if successful.
            ERROR_NOT_CONNECTED cannot connect to DB
            ERROR_MORE_DATA if buffer not large enough

--*)
function ComDBGetCurrentPortUsage(aHComDB: HCOMDB; var Buffer: Byte;
  BufferSize: DWORD; ReportType: ULONG; var MaxPortsReported: DWORD): LONG stdcall;


(*++

Routine Description:

    returns the first free COMx value

Arguments:

    Handle returned from OpenComPortDataBase.

Return Value:


    returns ERROR_SUCCESS if successful. or other ERROR_ if not

    if successful, then ComNumber will be that next free com value and claims it in the database


--*)
function ComDBClaimNextFreePort(aHComDB: HCOMDB; var ComNumber: DWORD): LONG stdcall;


(*++

Routine Description:

    Attempts to claim a com name in the database

Arguments:

    DataBaseHandle - returned from OpenComPortDataBase.

    ComNumber      - The port value to be claimed

    Force          - If TRUE, will force the port to be claimed even if in use already

    Forced         - will reflect the event that the claim was forced

Return Value:


    returns ERROR_SUCCESS if port name was not already claimed, or if it was claimed
                          and Force was TRUE.

            ERROR_SHARING_VIOLATION if port name is use and Force is false


--*)
function ComDBClaimPort(aHComDB: HCOMDB; ComNumber: DWORD; ForceClaim: BOOL;
  var Forced: BOOL): LONG stdcall;


(*++

Routine Description:

    Releases the port in the database

Arguments:

    DatabaseHandle - returned from OpenComPortDataBase.

    ComNumber      - port to be unclaimed in database

Return Value:


    returns ERROR_SUCCESS if successful
            ERROR_CANTWRITE if the changes cannot be committed
            ERROR_INVALID_PARAMETER if ComNumber is greater than the number of
                                    ports arbitrated


--*)
function ComDBReleasePort(aHComDB: HCOMDB; ComNumber: DWORD): LONG stdcall;


(*++

Routine Description:

    Resizes the database to the new size.  To get the current size, call
    ComDBGetCurrentPortUsage with a Buffer == NULL.

Arguments:

    DatabaseHandle - returned from OpenComPortDataBase.

    NewSize        - must be a multiple of 1024, with a max of 4096

Return Value:


    returns ERROR_SUCCESS if successful
            ERROR_CANTWRITE if the changes cannot be committed
            ERROR_BAD_LENGTH if NewSize is not greater than the current size or
                             NewSize is greater than COMDB_MAX_PORTS_ARBITRATED

--*)
function ComDBResizeDatabase(aHComDB: HCOMDB; NewSize: DWORD): LONG stdcall;

implementation

const
  MSPORTS_DLL = 'msports.dll';


{$IFNDEF SERIAL_STATIC_LINK}

  {$IFDEF SERIAL_USE_DELPHI_TYPES}

resourcestring
  SELibraryNotFound  = 'Library not found: %0:s';
  SEFunctionNotFound = 'Function not found: %0:s.%1:s';
  SEFunctionNotFound2 = 'Function not found: %0:s.%1:d';

procedure GetProcedureAddress(var P: Pointer; const ModuleName, ProcName: Ansistring);
  overload;
var
  ModuleHandle: HMODULE;
begin
  if not Assigned(P) then
  begin
    ModuleHandle := GetModuleHandle(PAnsiChar(Ansistring(ModuleName)));
    if ModuleHandle = 0 then
    begin
      ModuleHandle := LoadLibrary(PAnsiChar(ModuleName));
      if ModuleHandle = 0 then
        raise ELoadLibraryError.CreateFmt(SELibraryNotFound, [ModuleName]);
    end;
    P := Pointer(GetProcAddress(ModuleHandle, PAnsiChar(ProcName)));
    if not Assigned(P) then
      raise EGetProcAddressError.CreateFmt(SEFunctionNotFound, [ModuleName, ProcName]);
  end;
end;

procedure GetProcedureAddress(var P: Pointer; const ModuleName: Ansistring;
  ProcNumber: Cardinal); overload;
var
  ModuleHandle: HMODULE;
begin
  if not Assigned(P) then
  begin
    ModuleHandle := GetModuleHandle(PAnsiChar(Ansistring(ModuleName)));
    if ModuleHandle = 0 then
    begin
      ModuleHandle := LoadLibrary(PAnsiChar(ModuleName));
      if ModuleHandle = 0 then
        raise ELoadLibraryError.CreateFmt(SELibraryNotFound, [ModuleName]);
    end;
    P := Pointer(GetProcAddress(ModuleHandle, PAnsiChar(ProcNumber)));
    if not Assigned(P) then
      raise EGetProcAddressError.CreateFmt(SEFunctionNotFound2,
        [ModuleName, ProcNumber]);
  end;
end;

  {$ENDIF SERIAL_USE_DELPHI_TYPES}


  {$IFDEF SERIAL_ADVANCED_SETTINGS}

var
  _SerialDisplayAdvancedSettings: Pointer;

function SerialDisplayAdvancedSettings;
begin
  GetProcedureAddress(_SerialDisplayAdvancedSettings, MSPORTS_DLL,
    'SerialDisplayAdvancedSettings');
  asm
    MOV     ESP, EBP
    POP     EBP
    JMP     [_SerialDisplayAdvancedSettings]
  end;
end;

  {$ENDIF SERIAL_ADVANCED_SETTINGS}

var
  _ComDBOpen:  Pointer;
  _ComDBClose: Pointer;
  _ComDBGetCurrentPortUsage: Pointer;
  _ComDBClaimNextFreePort: Pointer;
  _ComDBClaimPort: Pointer;
  _ComDBReleasePort: Pointer;
  _ComDBResizeDatabase: Pointer;

function ComDBOpen;
begin
  GetProcedureAddress(_ComDBOpen, MSPORTS_DLL, 'ComDBOpen');
  asm
    MOV     ESP, EBP
    POP     EBP
    JMP     [_ComDBOpen]
  end;
end;

function ComDBClose;
begin
  GetProcedureAddress(_ComDBClose, MSPORTS_DLL, 'ComDBClose');
  asm
    MOV     ESP, EBP
    POP     EBP
    JMP     [_ComDBClose]
  end;
end;

function ComDBGetCurrentPortUsage;
begin
  GetProcedureAddress(_ComDBGetCurrentPortUsage, MSPORTS_DLL,
    'ComDBGetCurrentPortUsage');
  asm
    MOV     ESP, EBP
    POP     EBP
    JMP     [_ComDBGetCurrentPortUsage]
  end;
end;

function ComDBClaimNextFreePort;
begin
  GetProcedureAddress(_ComDBClaimNextFreePort, MSPORTS_DLL, 'ComDBClaimNextFreePort');
  asm
    MOV     ESP, EBP
    POP     EBP
    JMP     [_ComDBClaimNextFreePort]
  end;
end;

function ComDBClaimPort;
begin
  GetProcedureAddress(_ComDBClaimPort, MSPORTS_DLL, 'ComDBClaimPort');
  asm
    MOV     ESP, EBP
    POP     EBP
    JMP     [_ComDBClaimPort]
  end;
end;

function ComDBReleasePort;
begin
  GetProcedureAddress(_ComDBReleasePort, MSPORTS_DLL, 'ComDBReleasePort');
  asm
    MOV     ESP, EBP
    POP     EBP
    JMP     [_ComDBReleasePort]
  end;
end;

function ComDBResizeDatabase;
begin
  GetProcedureAddress(_ComDBResizeDatabase, MSPORTS_DLL, 'ComDBResizeDatabase');
  asm
    MOV     ESP, EBP
    POP     EBP
    JMP     [_ComDBResizeDatabase]
  end;
end;

{$ELSE SERIAL_STATIC_LINK}

  {$IFDEF SERIAL_ADVANCED_SETTINGS}
function SerialDisplayAdvancedSettings; external MSPORTS_DLL {$IFDEF DELAYED_LOADING} delayed{$ENDIF} Name 'SerialDisplayAdvancedSettings';
  {$ENDIF SERIAL_ADVANCED_SETTINGS}

function ComDBOpen; external MSPORTS_DLL {$IFDEF DELAYED_LOADING} delayed{$ENDIF} Name 'ComDBOpen';
function ComDBClose; external MSPORTS_DLL {$IFDEF DELAYED_LOADING} delayed{$ENDIF} Name 'ComDBClose';
function ComDBGetCurrentPortUsage;
  external MSPORTS_DLL {$IFDEF DELAYED_LOADING} delayed{$ENDIF} Name 'ComDBGetCurrentPortUsage';
function ComDBClaimNextFreePort;
  external MSPORTS_DLL {$IFDEF DELAYED_LOADING} delayed{$ENDIF} Name 'ComDBClaimNextFreePort';
function ComDBClaimPort; external MSPORTS_DLL
 {$IFDEF DELAYED_LOADING} delayed{$ENDIF} Name 'ComDBClaimPort';
function ComDBReleasePort;
  external MSPORTS_DLL {$IFDEF DELAYED_LOADING} delayed{$ENDIF} Name 'ComDBReleasePort';
function ComDBResizeDatabase;
  external MSPORTS_DLL {$IFDEF DELAYED_LOADING} delayed{$ENDIF} Name 'ComDBResizeDatabase';

{$ENDIF SERIAL_STATIC_LINK}

end.

