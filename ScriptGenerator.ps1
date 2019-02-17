#Script Generator
#2/17/2019 Austin Webber
#9/18/2018

#Resources Used: 
#MSI Information Grabber: http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/
#PS2EXE-GUI: https://gallery.technet.microsoft.com/scriptcenter/PS2EXE-GUI-Convert-e7cb69d5
#Create Shortcuts: https://www.pdq.com/blog/pdq-deploy-and-powershell/
#File Selection: https://blogs.technet.microsoft.com/heyscriptingguy/2009/09/01/hey-scripting-guy-can-i-open-a-file-dialog-box-with-windows-powershell/

#Converts PS to EXE
#Source https://gallery.technet.microsoft.com/scriptcenter/PS2EXE-GUI-Convert-e7cb69d5
function PS2EXE
{
	Param ([string]$inputFile = $null,
		[string]$outputFile = $null,
		[switch]$verbose,
		[switch]$debug,
		[switch]$runtime20,
		[switch]$runtime40,
		[switch]$x86,
		[switch]$x64,
		[int]$lcid,
		[switch]$Sta,
		[switch]$Mta,
		[switch]$noConsole,
		[switch]$nested,
		[string]$iconFile = $null,
		[string]$title,
		[string]$description,
		[string]$company,
		[string]$product,
		[string]$copyright,
		[string]$trademark,
		[string]$version,
		[switch]$requireAdmin,
		[switch]$virtualize,
		[switch]$credentialGUI,
		[switch]$noConfigfile)
	
<################################################################################>
<##                                                                            ##>
<##      PS2EXE-GUI v0.5.0.13                                                  ##>
<##      Written by: Ingo Karstein (http://blog.karstein-consulting.com)       ##>
<##      Reworked and GUI support by Markus Scholtes                           ##>
<##                                                                            ##>
<##      This script is released under Microsoft Public Licence                ##>
<##          that can be downloaded here:                                      ##>
<##          http://www.microsoft.com/opensource/licenses.mspx#Ms-PL           ##>
<##                                                                            ##>
<################################################################################>
	
	
	if (!$nested)
	{
		Write-Host "PS2EXE-GUI v0.5.0.13 by Ingo Karstein, reworked and GUI support by Markus Scholtes"
	}
	else
	{
		Write-Host "PowerShell 2.0 environment started..."
	}
	Write-Host ""
	
	if ($runtime20 -and $runtime40)
	{
		Write-Host "You cannot use switches -runtime20 and -runtime40 at the same time!"
		exit -1
	}
	
	if ($Sta -and $Mta)
	{
		Write-Host "You cannot use switches -Sta and -Mta at the same time!"
		exit -1
	}
	
	if ([string]::IsNullOrEmpty($inputFile) -or [string]::IsNullOrEmpty($outputFile))
	{
		Write-Host "Usage:"
		Write-Host ""
		Write-Host "powershell.exe -command ""&'.\ps2exe.ps1' [-inputFile] '<file_name>' [-outputFile] '<file_name>' [-verbose]"
		Write-Host "               [-debug] [-runtime20|-runtime40] [-lcid <id>] [-x86|-x64] [-Sta|-Mta] [-noConsole]"
		Write-Host "               [-credentialGUI] [-iconFile '<file_name>'] [-title '<title>'] [-description '<description>']"
		Write-Host "               [-company '<company>'] [-product '<product>'] [-copyright '<copyright>'] [-trademark '<trademark>']"
		Write-Host "               [-version '<version>'] [-noConfigfile] [-requireAdmin] [-virtualize]"""
		Write-Host ""
		Write-Host "    inputFile = Powershell script that you want to convert to EXE"
		Write-Host "   outputFile = destination EXE file name"
		Write-Host "      verbose = output verbose informations - if any"
		Write-Host "        debug = generate debug informations for output file"
		Write-Host "    runtime20 = this switch forces PS2EXE to create a config file for the generated EXE that contains the"
		Write-Host "                ""supported .NET Framework versions"" setting for .NET Framework 2.0/3.x for PowerShell 2.0"
		Write-Host "    runtime40 = this switch forces PS2EXE to create a config file for the generated EXE that contains the"
		Write-Host "                ""supported .NET Framework versions"" setting for .NET Framework 4.x for PowerShell 3.0 or higher"
		Write-Host "         lcid = location ID for the compiled EXE. Current user culture if not specified"
		Write-Host "          x86 = compile for 32-bit runtime only"
		Write-Host "          x64 = compile for 64-bit runtime only"
		Write-Host "          sta = Single Thread Apartment Mode"
		Write-Host "          mta = Multi Thread Apartment Mode"
		Write-Host "    noConsole = the resulting EXE file will be a Windows Forms app without a console window"
		Write-Host "credentialGUI = use GUI for prompting credentials in console mode"
		Write-Host "     iconFile = icon file name for the compiled EXE"
		Write-Host "        title = title information (displayed in details tab of Windows Explorer's properties dialog)"
		Write-Host "  description = description information (not displayed, but embedded in executable)"
		Write-Host "      company = company information (not displayed, but embedded in executable)"
		Write-Host "      product = product information (displayed in details tab of Windows Explorer's properties dialog)"
		Write-Host "    copyright = copyright information (displayed in details tab of Windows Explorer's properties dialog)"
		Write-Host "    trademark = trademark information (displayed in details tab of Windows Explorer's properties dialog)"
		Write-Host "      version = version information (displayed in details tab of Windows Explorer's properties dialog)"
		Write-Host " noConfigfile = write no config file (<outputfile>.exe.config)"
		Write-Host " requireAdmin = if UAC is enabled, compiled EXE run only in elevated context (UAC dialog appears if required)"
		Write-Host "   virtualize = application virtualization is activated (forcing x86 runtime)"
		Write-Host ""
		Write-Host "Input file or output file not specified!"
		exit -1
	}
	
	$psversion = 0
	if ($PSVersionTable.PSVersion.Major -ge 4)
	{
		$psversion = 4
		Write-Host "You are using PowerShell 4.0 or above."
	}
	
	if ($PSVersionTable.PSVersion.Major -eq 3)
	{
		$psversion = 3
		Write-Host "You are using PowerShell 3.0."
	}
	
	if ($PSVersionTable.PSVersion.Major -eq 2)
	{
		$psversion = 2
		Write-Host "You are using PowerShell 2.0."
	}
	
	if ($psversion -eq 0)
	{
		Write-Host "The powershell version is unknown!"
		exit -1
	}
	
	# retrieve absolute paths independent whether path is given relative oder absolute
	$inputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($inputFile)
	$outputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outputFile)
	
	if (!(Test-Path $inputFile -PathType Leaf))
	{
		Write-Host "Input file $($inputfile) not found!"
		exit -1
	}
	
	if ($inputFile -eq $outputFile)
	{
		Write-Host "Input file is identical to output file!"
		exit -1
	}
	
	if (!([string]::IsNullOrEmpty($iconFile)))
	{
		# retrieve absolute path independent whether path is given relative oder absolute
		$iconFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($iconFile)
		
		if (!(Test-Path $iconFile -PathType Leaf))
		{
			Write-Host "Icon file $($iconFile) not found!"
			exit -1
		}
	}
	
	if ($requireAdmin -And $virtualize)
	{
		Write-Host "-requireAdmin cannot be combined with -virtualize"
		exit -1
	}
	
	if (!$runtime20 -and !$runtime40)
	{
		if ($psversion -eq 4)
		{
			$runtime40 = $TRUE
		}
		elseif ($psversion -eq 3)
		{
			$runtime40 = $TRUE
		}
		else
		{
			$runtime20 = $TRUE
		}
	}
	
	if ($psversion -ge 3 -and $runtime20)
	{
		Write-Host "To create an EXE file for PowerShell 2.0 on PowerShell 3.0 or above this script now launches PowerShell 2.0..."
		Write-Host ""
		
		$arguments = "-inputFile '$($inputFile)' -outputFile '$($outputFile)' -nested "
		
		if ($verbose) { $arguments += "-verbose " }
		if ($debug) { $arguments += "-debug " }
		if ($runtime20) { $arguments += "-runtime20 " }
		if ($x86) { $arguments += "-x86 " }
		if ($x64) { $arguments += "-x64 " }
		if ($lcid) { $arguments += "-lcid $lcid " }
		if ($Sta) { $arguments += "-Sta " }
		if ($Mta) { $arguments += "-Mta " }
		if ($noConsole) { $arguments += "-noConsole " }
		if (!([string]::IsNullOrEmpty($iconFile))) { $arguments += "-iconFile '$($iconFile)' " }
		if (!([string]::IsNullOrEmpty($title))) { $arguments += "-title '$($title)' " }
		if (!([string]::IsNullOrEmpty($description))) { $arguments += "-description '$($description)' " }
		if (!([string]::IsNullOrEmpty($company))) { $arguments += "-company '$($company)' " }
		if (!([string]::IsNullOrEmpty($product))) { $arguments += "-product '$($product)' " }
		if (!([string]::IsNullOrEmpty($copyright))) { $arguments += "-copyright '$($copyright)' " }
		if (!([string]::IsNullOrEmpty($trademark))) { $arguments += "-trademark '$($trademark)' " }
		if (!([string]::IsNullOrEmpty($version))) { $arguments += "-version '$($version)' " }
		if ($requireAdmin) { $arguments += "-requireAdmin " }
		if ($virtualize) { $arguments += "-virtualize " }
		if ($credentialGUI) { $arguments += "-credentialGUI " }
		if ($noConfigfile) { $arguments += "-noConfigfile " }
		
		if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
		{
			# ps2exe.ps1 is running (script)
			$jobScript = @"
."$($PSHOME)\powershell.exe" -version 2.0 -command "&'$($MyInvocation.MyCommand.Path)' $($arguments)"
"@
		}
		else
		{
			# ps2exe.exe is running (compiled script)
			Write-Host "The parameter -runtime20 is not supported for compiled ps2exe.ps1 scripts."
			Write-Host "Compile ps2exe.ps1 with parameter -runtime20 and call the generated executable (without -runtime20)."
			exit -1
		}
		
		Invoke-Expression $jobScript
		
		exit 0
	}
	
	if ($psversion -lt 3 -and $runtime40)
	{
		Write-Host "You need to run ps2exe in an Powershell 3.0 or higher environment to use parameter -runtime40"
		Write-Host
		exit -1
	}
	
	if ($psversion -lt 3 -and !$Mta -and !$Sta)
	{
		# Set default apartment mode for powershell version if not set by parameter
		$Mta = $TRUE
	}
	
	if ($psversion -ge 3 -and !$Mta -and !$Sta)
	{
		# Set default apartment mode for powershell version if not set by parameter
		$Sta = $TRUE
	}
	
	# escape escape sequences in version info
	$title = $title -replace "\\", "\\"
	$product = $product -replace "\\", "\\"
	$copyright = $copyright -replace "\\", "\\"
	$trademark = $trademark -replace "\\", "\\"
	$description = $description -replace "\\", "\\"
	$company = $company -replace "\\", "\\"
	
	if (![string]::IsNullOrEmpty($version))
	{
		# check for correct version number information
		if ($version -notmatch "(^\d+\.\d+\.\d+\.\d+$)|(^\d+\.\d+\.\d+$)|(^\d+\.\d+$)|(^\d+$)")
		{
			Write-Host "Version number has to be supplied in the form n.n.n.n, n.n.n, n.n or n (with n as number)!"
			exit -1
		}
	}
	
	Write-Host ""
	
	$type = ('System.Collections.Generic.Dictionary`2') -as "Type"
	$type = $type.MakeGenericType(@(("System.String" -as "Type"), ("system.string" -as "Type")))
	$o = [Activator]::CreateInstance($type)
	
	$compiler20 = $FALSE
	if ($psversion -eq 3 -or $psversion -eq 4)
	{
		$o.Add("CompilerVersion", "v4.0")
	}
	else
	{
		if (Test-Path ("$ENV:WINDIR\Microsoft.NET\Framework\v3.5\csc.exe"))
		{ $o.Add("CompilerVersion", "v3.5") }
		else
		{
			Write-Warning "No .Net 3.5 compiler found, using .Net 2.0 compiler."
			Write-Warning "Therefore some methods are not available!"
			$compiler20 = $TRUE
			$o.Add("CompilerVersion", "v2.0")
		}
	}
	
	$referenceAssembies = @("System.dll")
	if (!$noConsole)
	{
		if ([System.AppDomain]::CurrentDomain.GetAssemblies() | ? { $_.ManifestModule.Name -ieq "Microsoft.PowerShell.ConsoleHost.dll" })
		{
			$referenceAssembies += ([System.AppDomain]::CurrentDomain.GetAssemblies() | ? { $_.ManifestModule.Name -ieq "Microsoft.PowerShell.ConsoleHost.dll" } | Select -First 1).Location
		}
	}
	$referenceAssembies += ([System.AppDomain]::CurrentDomain.GetAssemblies() | ? { $_.ManifestModule.Name -ieq "System.Management.Automation.dll" } | Select -First 1).Location
	
	if ($runtime40)
	{
		$n = New-Object System.Reflection.AssemblyName("System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
		[System.AppDomain]::CurrentDomain.Load($n) | Out-Null
		$referenceAssembies += ([System.AppDomain]::CurrentDomain.GetAssemblies() | ? { $_.ManifestModule.Name -ieq "System.Core.dll" } | Select -First 1).Location
	}
	
	if ($noConsole)
	{
		$n = New-Object System.Reflection.AssemblyName("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
		if ($runtime40)
		{
			$n = New-Object System.Reflection.AssemblyName("System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
		}
		[System.AppDomain]::CurrentDomain.Load($n) | Out-Null
		
		$n = New-Object System.Reflection.AssemblyName("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
		if ($runtime40)
		{
			$n = New-Object System.Reflection.AssemblyName("System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
		}
		[System.AppDomain]::CurrentDomain.Load($n) | Out-Null
		
		$referenceAssembies += ([System.AppDomain]::CurrentDomain.GetAssemblies() | ? { $_.ManifestModule.Name -ieq "System.Windows.Forms.dll" } | Select -First 1).Location
		$referenceAssembies += ([System.AppDomain]::CurrentDomain.GetAssemblies() | ? { $_.ManifestModule.Name -ieq "System.Drawing.dll" } | Select -First 1).Location
	}
	
	$platform = "anycpu"
	if ($x64 -and !$x86) { $platform = "x64" }
	else { if ($x86 -and !$x64) { $platform = "x86" } }
	
	$cop = (New-Object Microsoft.CSharp.CSharpCodeProvider($o))
	$cp = New-Object System.CodeDom.Compiler.CompilerParameters($referenceAssembies, $outputFile)
	$cp.GenerateInMemory = $FALSE
	$cp.GenerateExecutable = $TRUE
	
	$iconFileParam = ""
	if (!([string]::IsNullOrEmpty($iconFile)))
	{
		$iconFileParam = "`"/win32icon:$($iconFile)`""
	}
	
	$reqAdmParam = ""
	if ($requireAdmin)
	{
		$win32manifest = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>`r`n<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">`r`n<trustInfo xmlns=""urn:schemas-microsoft-com:asm.v2"">`r`n<security>`r`n<requestedPrivileges xmlns=""urn:schemas-microsoft-com:asm.v3"">`r`n<requestedExecutionLevel level=""requireAdministrator"" uiAccess=""false""/>`r`n</requestedPrivileges>`r`n</security>`r`n</trustInfo>`r`n</assembly>"
		$win32manifest | Set-Content ($outputFile + ".win32manifest") -Encoding UTF8
		
		$reqAdmParam = "`"/win32manifest:$($outputFile + ".win32manifest")`""
	}
	
	if (!$virtualize)
	{ $cp.CompilerOptions = "/platform:$($platform) /target:$(if ($noConsole) { 'winexe' }
			else { 'exe' }) $($iconFileParam) $($reqAdmParam)" }
	else
	{
		Write-Host "Application virtualization is activated, forcing x86 platfom."
		$cp.CompilerOptions = "/platform:x86 /target:$(if ($noConsole) { 'winexe' }
			else { 'exe' }) /nowin32manifest $($iconFileParam)"
	}
	
	$cp.IncludeDebugInformation = $debug
	
	if ($debug)
	{
		$cp.TempFiles.KeepFiles = $TRUE
	}
	
	Write-Host "Reading input file " -NoNewline
	Write-Host $inputFile
	Write-Host ""
	$content = Get-Content -LiteralPath ($inputFile) -Encoding UTF8 -ErrorAction SilentlyContinue
	if ($content -eq $null)
	{
		Write-Host "No data found. May be read error or file protected."
		exit -2
	}
	$scriptInp = [string]::Join("`r`n", $content)
	$script = [System.Convert]::ToBase64String(([System.Text.Encoding]::UTF8.GetBytes($scriptInp)))
	
	#region program frame
	$culture = ""
	
	if ($lcid)
	{
		$culture = @"
	System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo($lcid);
	System.Threading.Thread.CurrentThread.CurrentUICulture = System.Globalization.CultureInfo.GetCultureInfo($lcid);
"@
	}
	
	$programFrame = @"
// Simple PowerShell host created by Ingo Karstein (http://blog.karstein-consulting.com) for PS2EXE
// Reworked and GUI support by Markus Scholtes

using System;
using System.Collections.Generic;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using PowerShell = System.Management.Automation.PowerShell;
using System.Globalization;
using System.Management.Automation.Host;
using System.Security;
using System.Reflection;
using System.Runtime.InteropServices;
$(if ($noConsole) { @"
using System.Windows.Forms;
using System.Drawing;
"@
		})

[assembly:AssemblyTitle("$title")]
[assembly:AssemblyProduct("$product")]
[assembly:AssemblyCopyright("$copyright")]
[assembly:AssemblyTrademark("$trademark")]
$(if (![string]::IsNullOrEmpty($version)) { @"
[assembly:AssemblyVersion("$version")]
[assembly:AssemblyFileVersion("$version")]
"@
		})
// not displayed in details tab of properties dialog, but embedded to file
[assembly:AssemblyDescription("$description")]
[assembly:AssemblyCompany("$company")]

namespace ik.PowerShell
{
$(if ($noConsole -or $credentialGUI) { @"
	internal class CredentialForm
	{
		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
		private struct CREDUI_INFO
		{
			public int cbSize;
			public IntPtr hwndParent;
			public string pszMessageText;
			public string pszCaptionText;
			public IntPtr hbmBanner;
		}

		[Flags]
		enum CREDUI_FLAGS
		{
			INCORRECT_PASSWORD = 0x1,
			DO_NOT_PERSIST = 0x2,
			REQUEST_ADMINISTRATOR = 0x4,
			EXCLUDE_CERTIFICATES = 0x8,
			REQUIRE_CERTIFICATE = 0x10,
			SHOW_SAVE_CHECK_BOX = 0x40,
			ALWAYS_SHOW_UI = 0x80,
			REQUIRE_SMARTCARD = 0x100,
			PASSWORD_ONLY_OK = 0x200,
			VALIDATE_USERNAME = 0x400,
			COMPLETE_USERNAME = 0x800,
			PERSIST = 0x1000,
			SERVER_CREDENTIAL = 0x4000,
			EXPECT_CONFIRMATION = 0x20000,
			GENERIC_CREDENTIALS = 0x40000,
			USERNAME_TARGET_CREDENTIALS = 0x80000,
			KEEP_USERNAME = 0x100000,
		}

		public enum CredUIReturnCodes
		{
			NO_ERROR = 0,
			ERROR_CANCELLED = 1223,
			ERROR_NO_SUCH_LOGON_SESSION = 1312,
			ERROR_NOT_FOUND = 1168,
			ERROR_INVALID_ACCOUNT_NAME = 1315,
			ERROR_INSUFFICIENT_BUFFER = 122,
			ERROR_INVALID_PARAMETER = 87,
			ERROR_INVALID_FLAGS = 1004,
		}

		[DllImport("credui", CharSet = CharSet.Unicode)]
		private static extern CredUIReturnCodes CredUIPromptForCredentials(ref CREDUI_INFO creditUR,
			string targetName,
			IntPtr reserved1,
			int iError,
			StringBuilder userName,
			int maxUserName,
			StringBuilder password,
			int maxPassword,
			[MarshalAs(UnmanagedType.Bool)] ref bool pfSave,
			CREDUI_FLAGS flags);

		public class UserPwd
		{
			public string User = string.Empty;
			public string Password = string.Empty;
			public string Domain = string.Empty;
		}

		internal static UserPwd PromptForPassword(string caption, string message, string target, string user, PSCredentialTypes credTypes, PSCredentialUIOptions options)
		{
			// Flags und Variablen initialisieren
			StringBuilder userPassword = new StringBuilder(), userID = new StringBuilder(user, 128);
			CREDUI_INFO credUI = new CREDUI_INFO();
			if (!string.IsNullOrEmpty(message)) credUI.pszMessageText = message;
			if (!string.IsNullOrEmpty(caption)) credUI.pszCaptionText = caption;
			credUI.cbSize = Marshal.SizeOf(credUI);
			bool save = false;

			CREDUI_FLAGS flags = CREDUI_FLAGS.DO_NOT_PERSIST;
			if ((credTypes & PSCredentialTypes.Generic) == PSCredentialTypes.Generic)
			{
				flags |= CREDUI_FLAGS.GENERIC_CREDENTIALS;
				if ((options & PSCredentialUIOptions.AlwaysPrompt) == PSCredentialUIOptions.AlwaysPrompt)
				{
					flags |= CREDUI_FLAGS.ALWAYS_SHOW_UI;
				}
			}

			// den Benutzer nach Kennwort fragen, grafischer Prompt
			CredUIReturnCodes returnCode = CredUIPromptForCredentials(ref credUI, target, IntPtr.Zero, 0, userID, 128, userPassword, 128, ref save, flags);

			if (returnCode == CredUIReturnCodes.NO_ERROR)
			{
				UserPwd ret = new UserPwd();
				ret.User = userID.ToString();
				ret.Password = userPassword.ToString();
				ret.Domain = "";
				return ret;
			}

			return null;
		}
	}
"@
		})

	internal class PS2EXEHostRawUI : PSHostRawUserInterface
	{
$(if ($noConsole) { @"
		// Speicher fÃƒÂ¼r Konsolenfarben bei GUI-Output werden gelesen und gesetzt, aber im Moment nicht genutzt (for future use)
		private ConsoleColor ncBackgroundColor = ConsoleColor.White;
		private ConsoleColor ncForegroundColor = ConsoleColor.Black;
"@
		}
		else { @"
		const int STD_OUTPUT_HANDLE = -11;

		//CHAR_INFO struct, which was a union in the old days
		// so we want to use LayoutKind.Explicit to mimic it as closely
		// as we can
		[StructLayout(LayoutKind.Explicit)]
		public struct CHAR_INFO
		{
			[FieldOffset(0)]
			internal char UnicodeChar;
			[FieldOffset(0)]
			internal char AsciiChar;
			[FieldOffset(2)] //2 bytes seems to work properly
			internal UInt16 Attributes;
		}

		//COORD struct
		[StructLayout(LayoutKind.Sequential)]
		public struct COORD
		{
			public short X;
			public short Y;
		}

		//SMALL_RECT struct
		[StructLayout(LayoutKind.Sequential)]
		public struct SMALL_RECT
		{
			public short Left;
			public short Top;
			public short Right;
			public short Bottom;
		}

		/* Reads character and color attribute data from a rectangular block of character cells in a console screen buffer,
			 and the function writes the data to a rectangular block at a specified location in the destination buffer. */
		[DllImport("kernel32.dll", EntryPoint = "ReadConsoleOutputW", CharSet = CharSet.Unicode, SetLastError = true)]
		internal static extern bool ReadConsoleOutput(
			IntPtr hConsoleOutput,
			/* This pointer is treated as the origin of a two-dimensional array of CHAR_INFO structures
			whose size is specified by the dwBufferSize parameter.*/
			[MarshalAs(UnmanagedType.LPArray), Out] CHAR_INFO[,] lpBuffer,
			COORD dwBufferSize,
			COORD dwBufferCoord,
			ref SMALL_RECT lpReadRegion);

		/* Writes character and color attribute data to a specified rectangular block of character cells in a console screen buffer.
			The data to be written is taken from a correspondingly sized rectangular block at a specified location in the source buffer */
		[DllImport("kernel32.dll", EntryPoint = "WriteConsoleOutputW", CharSet = CharSet.Unicode, SetLastError = true)]
		internal static extern bool WriteConsoleOutput(
			IntPtr hConsoleOutput,
			/* This pointer is treated as the origin of a two-dimensional array of CHAR_INFO structures
			whose size is specified by the dwBufferSize parameter.*/
			[MarshalAs(UnmanagedType.LPArray), In] CHAR_INFO[,] lpBuffer,
			COORD dwBufferSize,
			COORD dwBufferCoord,
			ref SMALL_RECT lpWriteRegion);

		/* Moves a block of data in a screen buffer. The effects of the move can be limited by specifying a clipping rectangle, so
			the contents of the console screen buffer outside the clipping rectangle are unchanged. */
		[DllImport("kernel32.dll", SetLastError = true)]
		static extern bool ScrollConsoleScreenBuffer(
			IntPtr hConsoleOutput,
			[In] ref SMALL_RECT lpScrollRectangle,
			[In] ref SMALL_RECT lpClipRectangle,
			COORD dwDestinationOrigin,
			[In] ref CHAR_INFO lpFill);

		[DllImport("kernel32.dll", SetLastError = true)]
			static extern IntPtr GetStdHandle(int nStdHandle);
"@
		})

		public override ConsoleColor BackgroundColor
		{
$(if (!$noConsole) { @"
			get
			{
				return Console.BackgroundColor;
			}
			set
			{
				Console.BackgroundColor = value;
			}
"@
		}
		else { @"
			get
			{
				return ncBackgroundColor;
			}
			set
			{
				ncBackgroundColor = value;
			}
"@
		})
		}

		public override System.Management.Automation.Host.Size BufferSize
		{
			get
			{
$(if (!$noConsole) { @"
				if (ConsoleInfo.IsOutputRedirected())
					// return default value for redirection. If no valid value is returned WriteLine will not be called
					return new System.Management.Automation.Host.Size(120, 50);
				else
					return new System.Management.Automation.Host.Size(Console.BufferWidth, Console.BufferHeight);
"@
		}
		else { @"
					// return default value for Winforms. If no valid value is returned WriteLine will not be called
				return new System.Management.Automation.Host.Size(120, 50);
"@
		})
			}
			set
			{
$(if (!$noConsole) { @"
				Console.BufferWidth = value.Width;
				Console.BufferHeight = value.Height;
"@
		})
			}
		}

		public override Coordinates CursorPosition
		{
			get
			{
$(if (!$noConsole) { @"
				return new Coordinates(Console.CursorLeft, Console.CursorTop);
"@
		}
		else { @"
				// Dummywert fÃƒÂ¼r Winforms zurÃƒÂ¼ckgeben.
				return new Coordinates(0, 0);
"@
		})
			}
			set
			{
$(if (!$noConsole) { @"
				Console.CursorTop = value.Y;
				Console.CursorLeft = value.X;
"@
		})
			}
		}

		public override int CursorSize
		{
			get
			{
$(if (!$noConsole) { @"
				return Console.CursorSize;
"@
		}
		else { @"
				// Dummywert fÃƒÂ¼r Winforms zurÃƒÂ¼ckgeben.
				return 25;
"@
		})
			}
			set
			{
$(if (!$noConsole) { @"
				Console.CursorSize = value;
"@
		})
			}
		}

$(if ($noConsole) { @"
		private Form InvisibleForm = null;
"@
		})

		public override void FlushInputBuffer()
		{
$(if (!$noConsole) { @"
			if (!ConsoleInfo.IsInputRedirected())
			{	while (Console.KeyAvailable)
    			Console.ReadKey(true);
    	}
"@
		}
		else { @"
			if (InvisibleForm != null)
			{
				InvisibleForm.Close();
				InvisibleForm = null;
			}
			else
			{
				InvisibleForm = new Form();
				InvisibleForm.Opacity = 0;
				InvisibleForm.ShowInTaskbar = false;
				InvisibleForm.Visible = true;
			}
"@
		})
		}

		public override ConsoleColor ForegroundColor
		{
$(if (!$noConsole) { @"
			get
			{
				return Console.ForegroundColor;
			}
			set
			{
				Console.ForegroundColor = value;
			}
"@
		}
		else { @"
			get
			{
				return ncForegroundColor;
			}
			set
			{
				ncForegroundColor = value;
			}
"@
		})
		}

		public override BufferCell[,] GetBufferContents(System.Management.Automation.Host.Rectangle rectangle)
		{
$(if ($compiler20) { @"
			throw new Exception("Method GetBufferContents not implemented for .Net V2.0 compiler");
"@
		}
		else { if (!$noConsole) { @"
			IntPtr hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);
			CHAR_INFO[,] buffer = new CHAR_INFO[rectangle.Bottom - rectangle.Top + 1, rectangle.Right - rectangle.Left + 1];
			COORD buffer_size = new COORD() {X = (short)(rectangle.Right - rectangle.Left + 1), Y = (short)(rectangle.Bottom - rectangle.Top + 1)};
			COORD buffer_index = new COORD() {X = 0, Y = 0};
			SMALL_RECT screen_rect = new SMALL_RECT() {Left = (short)rectangle.Left, Top = (short)rectangle.Top, Right = (short)rectangle.Right, Bottom = (short)rectangle.Bottom};

			ReadConsoleOutput(hStdOut, buffer, buffer_size, buffer_index, ref screen_rect);

			System.Management.Automation.Host.BufferCell[,] ScreenBuffer = new System.Management.Automation.Host.BufferCell[rectangle.Bottom - rectangle.Top + 1, rectangle.Right - rectangle.Left + 1];
			for (int y = 0; y <= rectangle.Bottom - rectangle.Top; y++)
				for (int x = 0; x <= rectangle.Right - rectangle.Left; x++)
				{
					ScreenBuffer[y,x] = new System.Management.Automation.Host.BufferCell(buffer[y,x].AsciiChar, (System.ConsoleColor)(buffer[y,x].Attributes & 0xF), (System.ConsoleColor)((buffer[y,x].Attributes & 0xF0) / 0x10), System.Management.Automation.Host.BufferCellType.Complete);
				}

			return ScreenBuffer;
"@
			}
			else { @"
			System.Management.Automation.Host.BufferCell[,] ScreenBuffer = new System.Management.Automation.Host.BufferCell[rectangle.Bottom - rectangle.Top + 1, rectangle.Right - rectangle.Left + 1];

			for (int y = 0; y <= rectangle.Bottom - rectangle.Top; y++)
				for (int x = 0; x <= rectangle.Right - rectangle.Left; x++)
				{
					ScreenBuffer[y,x] = new System.Management.Automation.Host.BufferCell(' ', ncForegroundColor, ncBackgroundColor, System.Management.Automation.Host.BufferCellType.Complete);
				}

			return ScreenBuffer;
"@
			}
		})
		}

		public override bool KeyAvailable
		{
			get
			{
$(if (!$noConsole) { @"
				return Console.KeyAvailable;
"@
		}
		else { @"
				return true;
"@
		})
			}
		}

		public override System.Management.Automation.Host.Size MaxPhysicalWindowSize
		{
			get
			{
$(if (!$noConsole) { @"
				return new System.Management.Automation.Host.Size(Console.LargestWindowWidth, Console.LargestWindowHeight);
"@
		}
		else { @"
				// Dummy-Wert fÃƒÂ¼r Winforms
				return new System.Management.Automation.Host.Size(240, 84);
"@
		})
			}
		}

		public override System.Management.Automation.Host.Size MaxWindowSize
		{
			get
			{
$(if (!$noConsole) { @"
				return new System.Management.Automation.Host.Size(Console.BufferWidth, Console.BufferWidth);
"@
		}
		else { @"
				// Dummy-Wert fÃƒÂ¼r Winforms
				return new System.Management.Automation.Host.Size(120, 84);
"@
		})
			}
		}

		public override KeyInfo ReadKey(ReadKeyOptions options)
		{
$(if (!$noConsole) { @"
			ConsoleKeyInfo cki = Console.ReadKey((options & ReadKeyOptions.NoEcho)!=0);

			ControlKeyStates cks = 0;
			if ((cki.Modifiers & ConsoleModifiers.Alt) != 0)
				cks |= ControlKeyStates.LeftAltPressed | ControlKeyStates.RightAltPressed;
			if ((cki.Modifiers & ConsoleModifiers.Control) != 0)
				cks |= ControlKeyStates.LeftCtrlPressed | ControlKeyStates.RightCtrlPressed;
			if ((cki.Modifiers & ConsoleModifiers.Shift) != 0)
				cks |= ControlKeyStates.ShiftPressed;
			if (Console.CapsLock)
				cks |= ControlKeyStates.CapsLockOn;
			if (Console.NumberLock)
				cks |= ControlKeyStates.NumLockOn;

			return new KeyInfo((int)cki.Key, cki.KeyChar, cks, (options & ReadKeyOptions.IncludeKeyDown)!=0);
"@
		}
		else { @"
			if ((options & ReadKeyOptions.IncludeKeyDown)!=0)
				return ReadKeyBox.Show("", "", true);
			else
				return ReadKeyBox.Show("", "", false);
"@
		})
		}

		public override void ScrollBufferContents(System.Management.Automation.Host.Rectangle source, Coordinates destination, System.Management.Automation.Host.Rectangle clip, BufferCell fill)
		{ // no destination block clipping implemented
$(if (!$noConsole) { if ($compiler20) { @"
			throw new Exception("Method ScrollBufferContents not implemented for .Net V2.0 compiler");
"@
			}
			else { @"
			// clip area out of source range?
			if ((source.Left > clip.Right) || (source.Right < clip.Left) || (source.Top > clip.Bottom) || (source.Bottom < clip.Top))
			{ // clipping out of range -> nothing to do
				return;
			}

			IntPtr hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);
			SMALL_RECT lpScrollRectangle = new SMALL_RECT() {Left = (short)source.Left, Top = (short)source.Top, Right = (short)(source.Right), Bottom = (short)(source.Bottom)};
			SMALL_RECT lpClipRectangle;
			if (clip != null)
			{ lpClipRectangle = new SMALL_RECT() {Left = (short)clip.Left, Top = (short)clip.Top, Right = (short)(clip.Right), Bottom = (short)(clip.Bottom)}; }
			else
			{ lpClipRectangle = new SMALL_RECT() {Left = (short)0, Top = (short)0, Right = (short)(Console.WindowWidth - 1), Bottom = (short)(Console.WindowHeight - 1)}; }
			COORD dwDestinationOrigin = new COORD() {X = (short)(destination.X), Y = (short)(destination.Y)};
			CHAR_INFO lpFill = new CHAR_INFO() { AsciiChar = fill.Character, Attributes = (ushort)((int)(fill.ForegroundColor) + (int)(fill.BackgroundColor)*16) };

			ScrollConsoleScreenBuffer(hStdOut, ref lpScrollRectangle, ref lpClipRectangle, dwDestinationOrigin, ref lpFill);
"@
			}
		})
		}

		public override void SetBufferContents(System.Management.Automation.Host.Rectangle rectangle, BufferCell fill)
		{
$(if (!$noConsole) { @"
			// using a trick: move the buffer out of the screen, the source area gets filled with the char fill.Character
			if (rectangle.Left >= 0)
				Console.MoveBufferArea(rectangle.Left, rectangle.Top, rectangle.Right-rectangle.Left+1, rectangle.Bottom-rectangle.Top+1, BufferSize.Width, BufferSize.Height, fill.Character, fill.ForegroundColor, fill.BackgroundColor);
			else
			{ // Clear-Host: move all content off the screen
				Console.MoveBufferArea(0, 0, BufferSize.Width, BufferSize.Height, BufferSize.Width, BufferSize.Height, fill.Character, fill.ForegroundColor, fill.BackgroundColor);
			}
"@
		})
		}

		public override void SetBufferContents(Coordinates origin, BufferCell[,] contents)
		{
$(if (!$noConsole) { if ($compiler20) { @"
			throw new Exception("Method SetBufferContents not implemented for .Net V2.0 compiler");
"@
			}
			else { @"
			IntPtr hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);
			CHAR_INFO[,] buffer = new CHAR_INFO[contents.GetLength(0), contents.GetLength(1)];
			COORD buffer_size = new COORD() {X = (short)(contents.GetLength(1)), Y = (short)(contents.GetLength(0))};
			COORD buffer_index = new COORD() {X = 0, Y = 0};
			SMALL_RECT screen_rect = new SMALL_RECT() {Left = (short)origin.X, Top = (short)origin.Y, Right = (short)(origin.X + contents.GetLength(1) - 1), Bottom = (short)(origin.Y + contents.GetLength(0) - 1)};

			for (int y = 0; y < contents.GetLength(0); y++)
				for (int x = 0; x < contents.GetLength(1); x++)
				{
					buffer[y,x] = new CHAR_INFO() { AsciiChar = contents[y,x].Character, Attributes = (ushort)((int)(contents[y,x].ForegroundColor) + (int)(contents[y,x].BackgroundColor)*16) };
				}

			WriteConsoleOutput(hStdOut, buffer, buffer_size, buffer_index, ref screen_rect);
"@
			}
		})
		}

		public override Coordinates WindowPosition
		{
			get
			{
				Coordinates s = new Coordinates();
$(if (!$noConsole) { @"
				s.X = Console.WindowLeft;
				s.Y = Console.WindowTop;
"@
		}
		else { @"
				// Dummy-Wert fÃƒÂ¼r Winforms
				s.X = 0;
				s.Y = 0;
"@
		})
				return s;
			}
			set
			{
$(if (!$noConsole) { @"
				Console.WindowLeft = value.X;
				Console.WindowTop = value.Y;
"@
		})
			}
		}

		public override System.Management.Automation.Host.Size WindowSize
		{
			get
			{
				System.Management.Automation.Host.Size s = new System.Management.Automation.Host.Size();
$(if (!$noConsole) { @"
				s.Height = Console.WindowHeight;
				s.Width = Console.WindowWidth;
"@
		}
		else { @"
				// Dummy-Wert fÃƒÂ¼r Winforms
				s.Height = 50;
				s.Width = 120;
"@
		})
				return s;
			}
			set
			{
$(if (!$noConsole) { @"
				Console.WindowWidth = value.Width;
				Console.WindowHeight = value.Height;
"@
		})
			}
		}

		public override string WindowTitle
		{
			get
			{
$(if (!$noConsole) { @"
				return Console.Title;
"@
		}
		else { @"
				return System.AppDomain.CurrentDomain.FriendlyName;
"@
		})
			}
			set
			{
$(if (!$noConsole) { @"
				Console.Title = value;
"@
		})
			}
		}
	}

$(if ($noConsole) { @"
	public class InputBox
	{
		[DllImport("user32.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.Cdecl)]
		private static extern IntPtr MB_GetString(uint strId);

		public static DialogResult Show(string sTitle, string sPrompt, ref string sValue, bool bSecure)
		{
			// Generate controls
			Form form = new Form();
			Label label = new Label();
			TextBox textBox = new TextBox();
			Button buttonOk = new Button();
			Button buttonCancel = new Button();

			// Sizes and positions are defined according to the label
			// This control has to be finished first
			if (string.IsNullOrEmpty(sPrompt))
			{
				if (bSecure)
					label.Text = "Secure input:   ";
				else
					label.Text = "Input:          ";
			}
			else
				label.Text = sPrompt;
			label.Location = new Point(9, 19);
			label.AutoSize = true;
			// Size of the label is defined not before Add()
			form.Controls.Add(label);

			// Generate textbox
			if (bSecure) textBox.UseSystemPasswordChar = true;
			textBox.Text = sValue;
			textBox.SetBounds(12, label.Bottom, label.Right - 12, 20);

			// Generate buttons
			// get localized "OK"-string
			string sTextOK = Marshal.PtrToStringUni(MB_GetString(0));
			if (string.IsNullOrEmpty(sTextOK))
				buttonOk.Text = "OK";
			else
				buttonOk.Text = sTextOK;

			// get localized "Cancel"-string
			string sTextCancel = Marshal.PtrToStringUni(MB_GetString(1));
			if (string.IsNullOrEmpty(sTextCancel))
				buttonCancel.Text = "Cancel";
			else
				buttonCancel.Text = sTextCancel;

			buttonOk.DialogResult = DialogResult.OK;
			buttonCancel.DialogResult = DialogResult.Cancel;
			buttonOk.SetBounds(System.Math.Max(12, label.Right - 158), label.Bottom + 36, 75, 23);
			buttonCancel.SetBounds(System.Math.Max(93, label.Right - 77), label.Bottom + 36, 75, 23);

			// Configure form
			if (string.IsNullOrEmpty(sTitle))
				form.Text = System.AppDomain.CurrentDomain.FriendlyName;
			else
				form.Text = sTitle;
			form.ClientSize = new System.Drawing.Size(System.Math.Max(178, label.Right + 10), label.Bottom + 71);
			form.Controls.AddRange(new Control[] { textBox, buttonOk, buttonCancel });
			form.FormBorderStyle = FormBorderStyle.FixedDialog;
			form.StartPosition = FormStartPosition.CenterScreen;
			form.MinimizeBox = false;
			form.MaximizeBox = false;
			form.AcceptButton = buttonOk;
			form.CancelButton = buttonCancel;

			// Show form and compute results
			DialogResult dialogResult = form.ShowDialog();
			sValue = textBox.Text;
			return dialogResult;
		}

		public static DialogResult Show(string sTitle, string sPrompt, ref string sValue)
		{
			return Show(sTitle, sPrompt, ref sValue, false);
		}
	}

	public class ChoiceBox
	{
		public static int Show(System.Collections.ObjectModel.Collection<ChoiceDescription> aAuswahl, int iVorgabe, string sTitle, string sPrompt)
		{
			// cancel if array is empty
			if (aAuswahl == null) return -1;
			if (aAuswahl.Count < 1) return -1;

			// Generate controls
			Form form = new Form();
			RadioButton[] aradioButton = new RadioButton[aAuswahl.Count];
			ToolTip toolTip = new ToolTip();
			Button buttonOk = new Button();

			// Sizes and positions are defined according to the label
			// This control has to be finished first when a prompt is available
			int iPosY = 19, iMaxX = 0;
			if (!string.IsNullOrEmpty(sPrompt))
			{
				Label label = new Label();
				label.Text = sPrompt;
				label.Location = new Point(9, 19);
				label.AutoSize = true;
				// erst durch Add() wird die GrÃƒÂ¶ÃƒÅ¸e des Labels ermittelt
				form.Controls.Add(label);
				iPosY = label.Bottom;
				iMaxX = label.Right;
			}

			// An den Radiobuttons orientieren sich die weiteren GrÃƒÂ¶ÃƒÅ¸en und Positionen
			// Diese Controls also jetzt fertigstellen
			int Counter = 0;
			foreach (ChoiceDescription sAuswahl in aAuswahl)
			{
				aradioButton[Counter] = new RadioButton();
				aradioButton[Counter].Text = sAuswahl.Label;
				if (Counter == iVorgabe)
				{ aradioButton[Counter].Checked = true; }
				aradioButton[Counter].Location = new Point(9, iPosY);
				aradioButton[Counter].AutoSize = true;
				// erst durch Add() wird die GrÃƒÂ¶ÃƒÅ¸e des Labels ermittelt
				form.Controls.Add(aradioButton[Counter]);
				iPosY = aradioButton[Counter].Bottom;
				if (aradioButton[Counter].Right > iMaxX) { iMaxX = aradioButton[Counter].Right; }
				if (!string.IsNullOrEmpty(sAuswahl.HelpMessage))
				{
					 toolTip.SetToolTip(aradioButton[Counter], sAuswahl.HelpMessage);
				}
				Counter++;
			}

			// Tooltip auch anzeigen, wenn Parent-Fenster inaktiv ist
			toolTip.ShowAlways = true;

			// Button erzeugen
			buttonOk.Text = "OK";
			buttonOk.DialogResult = DialogResult.OK;
			buttonOk.SetBounds(System.Math.Max(12, iMaxX - 77), iPosY + 36, 75, 23);

			// configure form
			if (string.IsNullOrEmpty(sTitle))
				form.Text = System.AppDomain.CurrentDomain.FriendlyName;
			else
				form.Text = sTitle;
			form.ClientSize = new System.Drawing.Size(System.Math.Max(178, iMaxX + 10), iPosY + 71);
			form.Controls.Add(buttonOk);
			form.FormBorderStyle = FormBorderStyle.FixedDialog;
			form.StartPosition = FormStartPosition.CenterScreen;
			form.MinimizeBox = false;
			form.MaximizeBox = false;
			form.AcceptButton = buttonOk;

			// show and compute form
			if (form.ShowDialog() == DialogResult.OK)
			{ int iRueck = -1;
				for (Counter = 0; Counter < aAuswahl.Count; Counter++)
				{
					if (aradioButton[Counter].Checked == true)
					{ iRueck = Counter; }
				}
				return iRueck;
			}
			else
				return -1;
		}
	}

	public class ReadKeyBox
	{
		[DllImport("user32.dll")]
		public static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpKeyState,
			[Out, MarshalAs(UnmanagedType.LPWStr, SizeConst = 64)] System.Text.StringBuilder pwszBuff,
			int cchBuff, uint wFlags);

		static string GetCharFromKeys(Keys keys, bool bShift, bool bAltGr)
		{
			System.Text.StringBuilder buffer = new System.Text.StringBuilder(64);
			byte[] keyboardState = new byte[256];
			if (bShift)
			{ keyboardState[(int) Keys.ShiftKey] = 0xff; }
			if (bAltGr)
			{ keyboardState[(int) Keys.ControlKey] = 0xff;
				keyboardState[(int) Keys.Menu] = 0xff;
			}
			if (ToUnicode((uint) keys, 0, keyboardState, buffer, 64, 0) >= 1)
				return buffer.ToString();
			else
				return "\0";
		}

		class KeyboardForm : Form
		{
			public KeyboardForm()
			{
				this.KeyDown += new KeyEventHandler(KeyboardForm_KeyDown);
				this.KeyUp += new KeyEventHandler(KeyboardForm_KeyUp);
			}

			// check for KeyDown or KeyUp?
			public bool checkKeyDown = true;
			// key code for pressed key
			public KeyInfo keyinfo;

			void KeyboardForm_KeyDown(object sender, KeyEventArgs e)
			{
				if (checkKeyDown)
				{ // store key info
					keyinfo.VirtualKeyCode = e.KeyValue;
					keyinfo.Character = GetCharFromKeys(e.KeyCode, e.Shift, e.Alt & e.Control)[0];
					keyinfo.KeyDown = false;
					keyinfo.ControlKeyState = 0;
					if (e.Alt) { keyinfo.ControlKeyState = ControlKeyStates.LeftAltPressed | ControlKeyStates.RightAltPressed; }
					if (e.Control)
					{ keyinfo.ControlKeyState |= ControlKeyStates.LeftCtrlPressed | ControlKeyStates.RightCtrlPressed;
						if (!e.Alt)
						{ if (e.KeyValue > 64 && e.KeyValue < 96) keyinfo.Character = (char)(e.KeyValue - 64); }
					}
					if (e.Shift) { keyinfo.ControlKeyState |= ControlKeyStates.ShiftPressed; }
					if ((e.Modifiers & System.Windows.Forms.Keys.CapsLock) > 0) { keyinfo.ControlKeyState |= ControlKeyStates.CapsLockOn; }
					if ((e.Modifiers & System.Windows.Forms.Keys.NumLock) > 0) { keyinfo.ControlKeyState |= ControlKeyStates.NumLockOn; }
					// and close the form
					this.Close();
				}
			}

			void KeyboardForm_KeyUp(object sender, KeyEventArgs e)
			{
				if (!checkKeyDown)
				{ // store key info
					keyinfo.VirtualKeyCode = e.KeyValue;
					keyinfo.Character = GetCharFromKeys(e.KeyCode, e.Shift, e.Alt & e.Control)[0];
					keyinfo.KeyDown = true;
					keyinfo.ControlKeyState = 0;
					if (e.Alt) { keyinfo.ControlKeyState = ControlKeyStates.LeftAltPressed | ControlKeyStates.RightAltPressed; }
					if (e.Control)
					{ keyinfo.ControlKeyState |= ControlKeyStates.LeftCtrlPressed | ControlKeyStates.RightCtrlPressed;
						if (!e.Alt)
						{ if (e.KeyValue > 64 && e.KeyValue < 96) keyinfo.Character = (char)(e.KeyValue - 64); }
					}
					if (e.Shift) { keyinfo.ControlKeyState |= ControlKeyStates.ShiftPressed; }
					if ((e.Modifiers & System.Windows.Forms.Keys.CapsLock) > 0) { keyinfo.ControlKeyState |= ControlKeyStates.CapsLockOn; }
					if ((e.Modifiers & System.Windows.Forms.Keys.NumLock) > 0) { keyinfo.ControlKeyState |= ControlKeyStates.NumLockOn; }
					// and close the form
					this.Close();
				}
			}
		}

		public static KeyInfo Show(string sTitle, string sPrompt, bool bIncludeKeyDown)
		{
			// Controls erzeugen
			KeyboardForm form = new KeyboardForm();
			Label label = new Label();

			// Am Label orientieren sich die GrÃƒÂ¶ÃƒÅ¸en und Positionen
			// Dieses Control also zuerst fertigstellen
			if (string.IsNullOrEmpty(sPrompt))
			{
					label.Text = "Press a key";
			}
			else
				label.Text = sPrompt;
			label.Location = new Point(9, 19);
			label.AutoSize = true;
			// erst durch Add() wird die GrÃƒÂ¶ÃƒÅ¸e des Labels ermittelt
			form.Controls.Add(label);

			// configure form
			if (string.IsNullOrEmpty(sTitle))
				form.Text = System.AppDomain.CurrentDomain.FriendlyName;
			else
				form.Text = sTitle;
			form.ClientSize = new System.Drawing.Size(System.Math.Max(178, label.Right + 10), label.Bottom + 55);
			form.FormBorderStyle = FormBorderStyle.FixedDialog;
			form.StartPosition = FormStartPosition.CenterScreen;
			form.MinimizeBox = false;
			form.MaximizeBox = false;

			// show and compute form
			form.checkKeyDown = bIncludeKeyDown;
			form.ShowDialog();
			return form.keyinfo;
		}
	}

	public class ProgressForm : Form
	{
		private Label objLblActivity;
		private Label objLblStatus;
		private ProgressBar objProgressBar;
		private Label objLblRemainingTime;
		private Label objLblOperation;
		private ConsoleColor ProgressBarColor = ConsoleColor.DarkCyan;

		private Color DrawingColor(ConsoleColor color)
		{  // convert ConsoleColor to System.Drawing.Color
			switch (color)
			{
				case ConsoleColor.Black: return Color.Black;
				case ConsoleColor.Blue: return Color.Blue;
				case ConsoleColor.Cyan: return Color.Cyan;
				case ConsoleColor.DarkBlue: return ColorTranslator.FromHtml("#000080");
				case ConsoleColor.DarkGray: return ColorTranslator.FromHtml("#808080");
				case ConsoleColor.DarkGreen: return ColorTranslator.FromHtml("#008000");
				case ConsoleColor.DarkCyan: return ColorTranslator.FromHtml("#008080");
				case ConsoleColor.DarkMagenta: return ColorTranslator.FromHtml("#800080");
				case ConsoleColor.DarkRed: return ColorTranslator.FromHtml("#800000");
				case ConsoleColor.DarkYellow: return ColorTranslator.FromHtml("#808000");
				case ConsoleColor.Gray: return ColorTranslator.FromHtml("#C0C0C0");
				case ConsoleColor.Green: return ColorTranslator.FromHtml("#00FF00");
				case ConsoleColor.Magenta: return Color.Magenta;
				case ConsoleColor.Red: return Color.Red;
				case ConsoleColor.White: return Color.White;
				default: return Color.Yellow;
			}
		}

		private void InitializeComponent()
		{
			this.SuspendLayout();

			this.Text = "Progress";
			this.Height = 160;
			this.Width = 800;
			this.BackColor = Color.White;
			this.FormBorderStyle = FormBorderStyle.FixedSingle;
			this.ControlBox = false;
			this.StartPosition = FormStartPosition.CenterScreen;

			// Create Label
			objLblActivity = new Label();
			objLblActivity.Left = 5;
			objLblActivity.Top = 10;
			objLblActivity.Width = 800 - 20;
			objLblActivity.Height = 16;
			objLblActivity.Font = new Font(objLblActivity.Font, FontStyle.Bold);
			objLblActivity.Text = "";
			// Add Label to Form
			this.Controls.Add(objLblActivity);

			// Create Label
			objLblStatus = new Label();
			objLblStatus.Left = 25;
			objLblStatus.Top = 26;
			objLblStatus.Width = 800 - 40;
			objLblStatus.Height = 16;
			objLblStatus.Text = "";
			// Add Label to Form
			this.Controls.Add(objLblStatus);

			// Create ProgressBar
			objProgressBar = new ProgressBar();
			objProgressBar.Value = 0;
			objProgressBar.Style = ProgressBarStyle.Continuous;
			objProgressBar.ForeColor = DrawingColor(ProgressBarColor);
			objProgressBar.Size = new System.Drawing.Size(800 - 60, 20);
			objProgressBar.Left = 25;
			objProgressBar.Top = 55;
			// Add ProgressBar to Form
			this.Controls.Add(objProgressBar);

			// Create Label
			objLblRemainingTime = new Label();
			objLblRemainingTime.Left = 5;
			objLblRemainingTime.Top = 85;
			objLblRemainingTime.Width = 800 - 20;
			objLblRemainingTime.Height = 16;
			objLblRemainingTime.Text = "";
			// Add Label to Form
			this.Controls.Add(objLblRemainingTime);

			// Create Label
			objLblOperation = new Label();
			objLblOperation.Left = 25;
			objLblOperation.Top = 101;
			objLblOperation.Width = 800 - 40;
			objLblOperation.Height = 16;
			objLblOperation.Text = "";
			// Add Label to Form
			this.Controls.Add(objLblOperation);

			this.ResumeLayout();
		}

		public ProgressForm()
		{
			InitializeComponent();
		}

		public ProgressForm(ConsoleColor BarColor)
		{
			ProgressBarColor = BarColor;
			InitializeComponent();
		}

		public void Update(ProgressRecord objRecord)
		{
			if (objRecord == null)
				return;

			if (objRecord.RecordType == ProgressRecordType.Completed)
			{
				this.Close();
				return;
			}

			if (!string.IsNullOrEmpty(objRecord.Activity))
				objLblActivity.Text = objRecord.Activity;
			else
				objLblActivity.Text = "";

			if (!string.IsNullOrEmpty(objRecord.StatusDescription))
				objLblStatus.Text = objRecord.StatusDescription;
			else
				objLblStatus.Text = "";

			if ((objRecord.PercentComplete >= 0) && (objRecord.PercentComplete <= 100))
			{
				objProgressBar.Value = objRecord.PercentComplete;
				objProgressBar.Visible = true;
			}
			else
			{ if (objRecord.PercentComplete > 100)
				{
					objProgressBar.Value = 0;
					objProgressBar.Visible = true;
				}
				else
					objProgressBar.Visible = false;
			}

			if (objRecord.SecondsRemaining >= 0)
			{
				System.TimeSpan objTimeSpan = new System.TimeSpan(0, 0, objRecord.SecondsRemaining);
				objLblRemainingTime.Text = "Remaining time: " + string.Format("{0:00}:{1:00}:{2:00}", (int)objTimeSpan.TotalHours, objTimeSpan.Minutes, objTimeSpan.Seconds);
			}
			else
				objLblRemainingTime.Text = "";

			if (!string.IsNullOrEmpty(objRecord.CurrentOperation))
				objLblOperation.Text = objRecord.CurrentOperation;
			else
				objLblOperation.Text = "";

			this.Refresh();
			Application.DoEvents();
		}
	}
"@
		})

	// define IsInputRedirected(), IsOutputRedirected() and IsErrorRedirected() here since they were introduced first with .Net 4.5
	public class ConsoleInfo
	{
		private enum FileType : uint
		{
			FILE_TYPE_UNKNOWN = 0x0000,
			FILE_TYPE_DISK = 0x0001,
			FILE_TYPE_CHAR = 0x0002,
			FILE_TYPE_PIPE = 0x0003,
			FILE_TYPE_REMOTE = 0x8000
		}

		private enum STDHandle : uint
		{
			STD_INPUT_HANDLE = unchecked((uint)-10),
			STD_OUTPUT_HANDLE = unchecked((uint)-11),
			STD_ERROR_HANDLE = unchecked((uint)-12)
		}

		[DllImport("Kernel32.dll")]
		static private extern UIntPtr GetStdHandle(STDHandle stdHandle);

		[DllImport("Kernel32.dll")]
		static private extern FileType GetFileType(UIntPtr hFile);

		static public bool IsInputRedirected()
		{
			UIntPtr hInput = GetStdHandle(STDHandle.STD_INPUT_HANDLE);
			FileType fileType = (FileType)GetFileType(hInput);
			if ((fileType == FileType.FILE_TYPE_CHAR) || (fileType == FileType.FILE_TYPE_UNKNOWN))
				return false;
			return true;
		}

		static public bool IsOutputRedirected()
		{
			UIntPtr hOutput = GetStdHandle(STDHandle.STD_OUTPUT_HANDLE);
			FileType fileType = (FileType)GetFileType(hOutput);
			if ((fileType == FileType.FILE_TYPE_CHAR) || (fileType == FileType.FILE_TYPE_UNKNOWN))
				return false;
			return true;
		}

		static public bool IsErrorRedirected()
		{
			UIntPtr hError = GetStdHandle(STDHandle.STD_ERROR_HANDLE);
			FileType fileType = (FileType)GetFileType(hError);
			if ((fileType == FileType.FILE_TYPE_CHAR) || (fileType == FileType.FILE_TYPE_UNKNOWN))
				return false;
			return true;
		}
	}


	internal class PS2EXEHostUI : PSHostUserInterface
	{
		private PS2EXEHostRawUI rawUI = null;

		public ConsoleColor ErrorForegroundColor = ConsoleColor.Red;
		public ConsoleColor ErrorBackgroundColor = ConsoleColor.Black;

		public ConsoleColor WarningForegroundColor = ConsoleColor.Yellow;
		public ConsoleColor WarningBackgroundColor = ConsoleColor.Black;

		public ConsoleColor DebugForegroundColor = ConsoleColor.Yellow;
		public ConsoleColor DebugBackgroundColor = ConsoleColor.Black;

		public ConsoleColor VerboseForegroundColor = ConsoleColor.Yellow;
		public ConsoleColor VerboseBackgroundColor = ConsoleColor.Black;

$(if (!$noConsole) { @"
		public ConsoleColor ProgressForegroundColor = ConsoleColor.Yellow;
"@
		}
		else { @"
		public ConsoleColor ProgressForegroundColor = ConsoleColor.DarkCyan;
"@
		})
		public ConsoleColor ProgressBackgroundColor = ConsoleColor.DarkCyan;

		public PS2EXEHostUI() : base()
		{
			rawUI = new PS2EXEHostRawUI();
$(if (!$noConsole) { @"
			rawUI.ForegroundColor = Console.ForegroundColor;
			rawUI.BackgroundColor = Console.BackgroundColor;
"@
		})
		}

		public override Dictionary<string, PSObject> Prompt(string caption, string message, System.Collections.ObjectModel.Collection<FieldDescription> descriptions)
		{
$(if (!$noConsole) { @"
			if (!string.IsNullOrEmpty(caption)) WriteLine(caption);
			if (!string.IsNullOrEmpty(message)) WriteLine(message);
"@
		}
		else { @"
			if ((!string.IsNullOrEmpty(caption)) || (!string.IsNullOrEmpty(message)))
			{ string sTitel = System.AppDomain.CurrentDomain.FriendlyName, sMeldung = "";

				if (!string.IsNullOrEmpty(caption)) sTitel = caption;
				if (!string.IsNullOrEmpty(message)) sMeldung = message;
				MessageBox.Show(sMeldung, sTitel);
			}

			// Titel und Labeltext fÃƒÂ¼r Inputbox zurÃƒÂ¼cksetzen
			ibcaption = "";
			ibmessage = "";
"@
		})
			Dictionary<string, PSObject> ret = new Dictionary<string, PSObject>();
			foreach (FieldDescription cd in descriptions)
			{
				Type t = null;
				if (string.IsNullOrEmpty(cd.ParameterAssemblyFullName))
					t = typeof(string);
				else
					t = Type.GetType(cd.ParameterAssemblyFullName);

				if (t.IsArray)
				{
					Type elementType = t.GetElementType();
					Type genericListType = Type.GetType("System.Collections.Generic.List"+((char)0x60).ToString()+"1");
					genericListType = genericListType.MakeGenericType(new Type[] { elementType });
					ConstructorInfo constructor = genericListType.GetConstructor(BindingFlags.CreateInstance | BindingFlags.Instance | BindingFlags.Public, null, Type.EmptyTypes, null);
					object resultList = constructor.Invoke(null);

					int index = 0;
					string data = "";
					do
					{
						try
						{
$(if (!$noConsole) { @"
							if (!string.IsNullOrEmpty(cd.Name)) Write(string.Format("{0}[{1}]: ", cd.Name, index));
"@
		}
		else { @"
							if (!string.IsNullOrEmpty(cd.Name)) ibmessage = string.Format("{0}[{1}]: ", cd.Name, index);
"@
		})
							data = ReadLine();
							if (string.IsNullOrEmpty(data))
								break;

							object o = System.Convert.ChangeType(data, elementType);
							genericListType.InvokeMember("Add", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, resultList, new object[] { o });
						}
						catch (Exception e)
						{
							throw e;
						}
						index++;
					} while (true);

					System.Array retArray = (System.Array )genericListType.InvokeMember("ToArray", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, resultList, null);
					ret.Add(cd.Name, new PSObject(retArray));
				}
				else
				{
					object o = null;
					string l = null;
					try
					{
						if (t != typeof(System.Security.SecureString))
						{
							if (t != typeof(System.Management.Automation.PSCredential))
							{
$(if (!$noConsole) { @"
								if (!string.IsNullOrEmpty(cd.Name)) Write(cd.Name);
								if (!string.IsNullOrEmpty(cd.HelpMessage)) Write(" (Type !? for help.)");
								if ((!string.IsNullOrEmpty(cd.Name)) || (!string.IsNullOrEmpty(cd.HelpMessage))) Write(": ");
"@
		}
		else { @"
								if (!string.IsNullOrEmpty(cd.Name)) ibmessage = string.Format("{0}: ", cd.Name);
								if (!string.IsNullOrEmpty(cd.HelpMessage)) ibmessage += "\n(Type !? for help.)";
"@
		})
								do {
									l = ReadLine();
									if (l == "!?")
										WriteLine(cd.HelpMessage);
									else
									{
										if (string.IsNullOrEmpty(l)) o = cd.DefaultValue;
										if (o == null)
										{
											try {
												o = System.Convert.ChangeType(l, t);
											}
											catch {
												Write("Wrong format, please repeat input: ");
												l = "!?";
											}
										}
									}
								} while (l == "!?");
							}
							else
							{
								PSCredential pscred = PromptForCredential("", "", "", "");
								o = pscred;
							}
						}
						else
						{
$(if (!$noConsole) { @"
								if (!string.IsNullOrEmpty(cd.Name)) Write(string.Format("{0}: ", cd.Name));
"@
		}
		else { @"
								if (!string.IsNullOrEmpty(cd.Name)) ibmessage = string.Format("{0}: ", cd.Name);
"@
		})

							SecureString pwd = null;
							pwd = ReadLineAsSecureString();
							o = pwd;
						}

						ret.Add(cd.Name, new PSObject(o));
					}
					catch (Exception e)
					{
						throw e;
					}
				}
			}
$(if ($noConsole) { @"
			// Titel und Labeltext fÃƒÂ¼r Inputbox zurÃƒÂ¼cksetzen
			ibcaption = "";
			ibmessage = "";
"@
		})
			return ret;
		}

		public override int PromptForChoice(string caption, string message, System.Collections.ObjectModel.Collection<ChoiceDescription> choices, int defaultChoice)
		{
$(if ($noConsole) { @"
			int iReturn = ChoiceBox.Show(choices, defaultChoice, caption, message);
			if (iReturn == -1) { iReturn = defaultChoice; }
			return iReturn;
"@
		}
		else { @"
			if (!string.IsNullOrEmpty(caption))
				WriteLine(caption);
			WriteLine(message);
			int idx = 0;
			SortedList<string, int> res = new SortedList<string, int>();
			foreach (ChoiceDescription cd in choices)
			{
				string lkey = cd.Label.Substring(0, 1), ltext = cd.Label;
				int pos = cd.Label.IndexOf('&');
				if (pos > -1)
				{
					lkey = cd.Label.Substring(pos + 1, 1).ToUpper();
					if (pos > 0)
						ltext = cd.Label.Substring(0, pos) + cd.Label.Substring(pos + 1);
					else
						ltext = cd.Label.Substring(1);
				}
				res.Add(lkey.ToLower(), idx);

				if (idx > 0) Write("  ");
				if (idx == defaultChoice)
				{
					Write(ConsoleColor.Yellow, Console.BackgroundColor, string.Format("[{0}] {1}", lkey, ltext));
					if (!string.IsNullOrEmpty(cd.HelpMessage))
						Write(ConsoleColor.Gray, Console.BackgroundColor, string.Format(" ({0})", cd.HelpMessage));
				}
				else
				{
					Write(ConsoleColor.Gray, Console.BackgroundColor, string.Format("[{0}] {1}", lkey, ltext));
					if (!string.IsNullOrEmpty(cd.HelpMessage))
						Write(ConsoleColor.Gray, Console.BackgroundColor, string.Format(" ({0})", cd.HelpMessage));
				}
				idx++;
			}
			Write(": ");

			try
			{
				while (true)
				{ string s = Console.ReadLine().ToLower();
					if (res.ContainsKey(s))
						return res[s];
					if (string.IsNullOrEmpty(s))
						return defaultChoice;
				}
			}
			catch { }

			return defaultChoice;
"@
		})
		}

		public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName, PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options)
		{
$(if (!$noConsole -and !$credentialGUI) { @"
			if (!string.IsNullOrEmpty(caption)) WriteLine(caption);
			WriteLine(message);

			string un;
			if ((string.IsNullOrEmpty(userName)) || ((options & PSCredentialUIOptions.ReadOnlyUserName) == 0))
			{
				Write("User name: ");
				un = ReadLine();
			}
			else
			{
				Write("User name: ");
				if (!string.IsNullOrEmpty(targetName)) Write(targetName + "\\");
				WriteLine(userName);
				un = userName;
			}
			SecureString pwd = null;
			Write("Password: ");
			pwd = ReadLineAsSecureString();

			if (string.IsNullOrEmpty(un)) un = "<NOUSER>";
			if (!string.IsNullOrEmpty(targetName))
			{
				if (un.IndexOf('\\') < 0)
					un = targetName + "\\" + un;
			}

			PSCredential c2 = new PSCredential(un, pwd);
			return c2;
"@
		}
		else { @"
			ik.PowerShell.CredentialForm.UserPwd cred = CredentialForm.PromptForPassword(caption, message, targetName, userName, allowedCredentialTypes, options);
			if (cred != null)
			{
				System.Security.SecureString x = new System.Security.SecureString();
				foreach (char c in cred.Password.ToCharArray())
					x.AppendChar(c);

				return new PSCredential(cred.User, x);
			}
			return null;
"@
		})
		}

		public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName)
		{
$(if (!$noConsole -and !$credentialGUI) { @"
			if (!string.IsNullOrEmpty(caption)) WriteLine(caption);
			WriteLine(message);

			string un;
			if (string.IsNullOrEmpty(userName))
			{
				Write("User name: ");
				un = ReadLine();
			}
			else
			{
				Write("User name: ");
				if (!string.IsNullOrEmpty(targetName)) Write(targetName + "\\");
				WriteLine(userName);
				un = userName;
			}
			SecureString pwd = null;
			Write("Password: ");
			pwd = ReadLineAsSecureString();

			if (string.IsNullOrEmpty(un)) un = "<NOUSER>";
			if (!string.IsNullOrEmpty(targetName))
			{
				if (un.IndexOf('\\') < 0)
					un = targetName + "\\" + un;
			}

			PSCredential c2 = new PSCredential(un, pwd);
			return c2;
"@
		}
		else { @"
			ik.PowerShell.CredentialForm.UserPwd cred = CredentialForm.PromptForPassword(caption, message, targetName, userName, PSCredentialTypes.Default, PSCredentialUIOptions.Default);
			if (cred != null)
			{
				System.Security.SecureString x = new System.Security.SecureString();
				foreach (char c in cred.Password.ToCharArray())
					x.AppendChar(c);

				return new PSCredential(cred.User, x);
			}
			return null;
"@
		})
		}

		public override PSHostRawUserInterface RawUI
		{
			get
			{
				return rawUI;
			}
		}

$(if ($noConsole) { @"
		private string ibcaption;
		private string ibmessage;
"@
		})

		public override string ReadLine()
		{
$(if (!$noConsole) { @"
			return Console.ReadLine();
"@
		}
		else { @"
			string sWert = "";
			if (InputBox.Show(ibcaption, ibmessage, ref sWert) == DialogResult.OK)
				return sWert;
			else
				return "";
"@
		})
		}

		private System.Security.SecureString getPassword()
		{
			System.Security.SecureString pwd = new System.Security.SecureString();
			while (true)
			{
				ConsoleKeyInfo i = Console.ReadKey(true);
				if (i.Key == ConsoleKey.Enter)
				{
					Console.WriteLine();
					break;
				}
				else if (i.Key == ConsoleKey.Backspace)
				{
					if (pwd.Length > 0)
					{
						pwd.RemoveAt(pwd.Length - 1);
						Console.Write("\b \b");
					}
				}
				else
				{
					pwd.AppendChar(i.KeyChar);
					Console.Write("*");
				}
			}
			return pwd;
		}

		public override System.Security.SecureString ReadLineAsSecureString()
		{
			System.Security.SecureString secstr = new System.Security.SecureString();
$(if (!$noConsole) { @"
			secstr = getPassword();
"@
		}
		else { @"
			string sWert = "";

			if (InputBox.Show(ibcaption, ibmessage, ref sWert, true) == DialogResult.OK)
			{
				foreach (char ch in sWert)
					secstr.AppendChar(ch);
			}
"@
		})
			return secstr;
		}

		// called by Write-Host
		public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value)
		{
$(if (!$noConsole) { @"
			ConsoleColor fgc = Console.ForegroundColor, bgc = Console.BackgroundColor;
			Console.ForegroundColor = foregroundColor;
			Console.BackgroundColor = backgroundColor;
			Console.Write(value);
			Console.ForegroundColor = fgc;
			Console.BackgroundColor = bgc;
"@
		}
		else { @"
			if ((!string.IsNullOrEmpty(value)) && (value != "\n"))
				MessageBox.Show(value, System.AppDomain.CurrentDomain.FriendlyName);
"@
		})
		}

		public override void Write(string value)
		{
$(if (!$noConsole) { @"
			Console.Write(value);
"@
		}
		else { @"
			if ((!string.IsNullOrEmpty(value)) && (value != "\n"))
				MessageBox.Show(value, System.AppDomain.CurrentDomain.FriendlyName);
"@
		})
		}

		// called by Write-Debug
		public override void WriteDebugLine(string message)
		{
$(if (!$noConsole) { @"
			WriteLine(DebugForegroundColor, DebugBackgroundColor, string.Format("DEBUG: {0}", message));
"@
		}
		else { @"
			MessageBox.Show(message, System.AppDomain.CurrentDomain.FriendlyName, MessageBoxButtons.OK, MessageBoxIcon.Information);
"@
		})
		}

		// called by Write-Error
		public override void WriteErrorLine(string value)
		{
$(if (!$noConsole) { @"
			if (ConsoleInfo.IsErrorRedirected())
				Console.Error.WriteLine(string.Format("ERROR: {0}", value));
			else
				WriteLine(ErrorForegroundColor, ErrorBackgroundColor, string.Format("ERROR: {0}", value));
"@
		}
		else { @"
			MessageBox.Show(value, System.AppDomain.CurrentDomain.FriendlyName, MessageBoxButtons.OK, MessageBoxIcon.Error);
"@
		})
		}

		public override void WriteLine()
		{
$(if (!$noConsole) { @"
			Console.WriteLine();
"@
		}
		else { @"
			MessageBox.Show("", System.AppDomain.CurrentDomain.FriendlyName);
"@
		})
		}

		public override void WriteLine(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value)
		{
$(if (!$noConsole) { @"
			ConsoleColor fgc = Console.ForegroundColor, bgc = Console.BackgroundColor;
			Console.ForegroundColor = foregroundColor;
			Console.BackgroundColor = backgroundColor;
			Console.WriteLine(value);
			Console.ForegroundColor = fgc;
			Console.BackgroundColor = bgc;
"@
		}
		else { @"
			if ((!string.IsNullOrEmpty(value)) && (value != "\n"))
				MessageBox.Show(value, System.AppDomain.CurrentDomain.FriendlyName);
"@
		})
		}

		// called by Write-Output
		public override void WriteLine(string value)
		{
$(if (!$noConsole) { @"
			Console.WriteLine(value);
"@
		}
		else { @"
			if ((!string.IsNullOrEmpty(value)) && (value != "\n"))
				MessageBox.Show(value, System.AppDomain.CurrentDomain.FriendlyName);
"@
		})
		}

$(if ($noConsole) { @"
		public ProgressForm pf = null;
"@
		})
		public override void WriteProgress(long sourceId, ProgressRecord record)
		{
$(if ($noConsole) { @"
			if (pf == null)
			{
				pf = new ProgressForm(ProgressForegroundColor);
				pf.Show();
			}
			pf.Update(record);
			if (record.RecordType == ProgressRecordType.Completed)
			{
				pf = null;
			}
"@
		})
		}

		// called by Write-Verbose
		public override void WriteVerboseLine(string message)
		{
$(if (!$noConsole) { @"
			WriteLine(VerboseForegroundColor, VerboseBackgroundColor, string.Format("VERBOSE: {0}", message));
"@
		}
		else { @"
			MessageBox.Show(message, System.AppDomain.CurrentDomain.FriendlyName, MessageBoxButtons.OK, MessageBoxIcon.Information);
"@
		})
		}

		// called by Write-Warning
		public override void WriteWarningLine(string message)
		{
$(if (!$noConsole) { @"
			WriteLine(WarningForegroundColor, WarningBackgroundColor, string.Format("WARNING: {0}", message));
"@
		}
		else { @"
			MessageBox.Show(message, System.AppDomain.CurrentDomain.FriendlyName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
"@
		})
		}
	}

	internal class PS2EXEHost : PSHost
	{
		private PS2EXEApp parent;
		private PS2EXEHostUI ui = null;

		private CultureInfo originalCultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;

		private CultureInfo originalUICultureInfo = System.Threading.Thread.CurrentThread.CurrentUICulture;

		private Guid myId = Guid.NewGuid();

		public PS2EXEHost(PS2EXEApp app, PS2EXEHostUI ui)
		{
			this.parent = app;
			this.ui = ui;
		}

		public class ConsoleColorProxy
		{
			private PS2EXEHostUI _ui;

			public ConsoleColorProxy(PS2EXEHostUI ui)
			{
				if (ui == null) throw new ArgumentNullException("ui");
				_ui = ui;
			}

			public ConsoleColor ErrorForegroundColor
			{
				get
				{ return _ui.ErrorForegroundColor; }
				set
				{ _ui.ErrorForegroundColor = value; }
			}

			public ConsoleColor ErrorBackgroundColor
			{
				get
				{ return _ui.ErrorBackgroundColor; }
				set
				{ _ui.ErrorBackgroundColor = value; }
			}

			public ConsoleColor WarningForegroundColor
			{
				get
				{ return _ui.WarningForegroundColor; }
				set
				{ _ui.WarningForegroundColor = value; }
			}

			public ConsoleColor WarningBackgroundColor
			{
				get
				{ return _ui.WarningBackgroundColor; }
				set
				{ _ui.WarningBackgroundColor = value; }
			}

			public ConsoleColor DebugForegroundColor
			{
				get
				{ return _ui.DebugForegroundColor; }
				set
				{ _ui.DebugForegroundColor = value; }
			}

			public ConsoleColor DebugBackgroundColor
			{
				get
				{ return _ui.DebugBackgroundColor; }
				set
				{ _ui.DebugBackgroundColor = value; }
			}

			public ConsoleColor VerboseForegroundColor
			{
				get
				{ return _ui.VerboseForegroundColor; }
				set
				{ _ui.VerboseForegroundColor = value; }
			}

			public ConsoleColor VerboseBackgroundColor
			{
				get
				{ return _ui.VerboseBackgroundColor; }
				set
				{ _ui.VerboseBackgroundColor = value; }
			}

			public ConsoleColor ProgressForegroundColor
			{
				get
				{ return _ui.ProgressForegroundColor; }
				set
				{ _ui.ProgressForegroundColor = value; }
			}

			public ConsoleColor ProgressBackgroundColor
			{
				get
				{ return _ui.ProgressBackgroundColor; }
				set
				{ _ui.ProgressBackgroundColor = value; }
			}
		}

		public override PSObject PrivateData
		{
			get
			{
				if (ui == null) return null;
				return _consoleColorProxy ?? (_consoleColorProxy = PSObject.AsPSObject(new ConsoleColorProxy(ui)));
			}
		}

		private PSObject _consoleColorProxy;

		public override System.Globalization.CultureInfo CurrentCulture
		{
			get
			{
				return this.originalCultureInfo;
			}
		}

		public override System.Globalization.CultureInfo CurrentUICulture
		{
			get
			{
				return this.originalUICultureInfo;
			}
		}

		public override Guid InstanceId
		{
			get
			{
				return this.myId;
			}
		}

		public override string Name
		{
			get
			{
				return "PS2EXE_Host";
			}
		}

		public override PSHostUserInterface UI
		{
			get
			{
				return ui;
			}
		}

		public override Version Version
		{
			get
			{
				return new Version(0, 5, 0, 13);
			}
		}

		public override void EnterNestedPrompt()
		{
		}

		public override void ExitNestedPrompt()
		{
		}

		public override void NotifyBeginApplication()
		{
			return;
		}

		public override void NotifyEndApplication()
		{
			return;
		}

		public override void SetShouldExit(int exitCode)
		{
			this.parent.ShouldExit = true;
			this.parent.ExitCode = exitCode;
		}
	}

	internal interface PS2EXEApp
	{
		bool ShouldExit { get; set; }
		int ExitCode { get; set; }
	}

	internal class PS2EXE : PS2EXEApp
	{
		private bool shouldExit;

		private int exitCode;

		public bool ShouldExit
		{
			get { return this.shouldExit; }
			set { this.shouldExit = value; }
		}

		public int ExitCode
		{
			get { return this.exitCode; }
			set { this.exitCode = value; }
		}

		$(if ($Sta) { "[STAThread]" })$(if ($Mta) { "[MTAThread]" })
		private static int Main(string[] args)
		{
			$culture

			PS2EXE me = new PS2EXE();

			bool paramWait = false;
			string extractFN = string.Empty;

			PS2EXEHostUI ui = new PS2EXEHostUI();
			PS2EXEHost host = new PS2EXEHost(me, ui);
			System.Threading.ManualResetEvent mre = new System.Threading.ManualResetEvent(false);

			AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

			try
			{
				using (Runspace myRunSpace = RunspaceFactory.CreateRunspace(host))
				{
					$(if ($Sta -or $Mta) { "myRunSpace.ApartmentState = System.Threading.ApartmentState." })$(if ($Sta) { "STA" })$(if ($Mta) { "MTA" });
					myRunSpace.Open();

					using (System.Management.Automation.PowerShell powershell = System.Management.Automation.PowerShell.Create())
					{
$(if (!$noConsole) { @"
						Console.CancelKeyPress += new ConsoleCancelEventHandler(delegate(object sender, ConsoleCancelEventArgs e)
						{
							try
							{
								powershell.BeginStop(new AsyncCallback(delegate(IAsyncResult r)
								{
									mre.Set();
									e.Cancel = true;
								}), null);
							}
							catch
							{
							};
						});
"@
		})

						powershell.Runspace = myRunSpace;
						powershell.Streams.Error.DataAdded += new EventHandler<DataAddedEventArgs>(delegate(object sender, DataAddedEventArgs e)
						{
							ui.WriteErrorLine(((PSDataCollection<ErrorRecord>)sender)[e.Index].ToString());
						});

						PSDataCollection<string> colInput = new PSDataCollection<string>();
$(if (!$runtime20) { @"
						if (ConsoleInfo.IsInputRedirected())
						{ // read standard input
							string sItem = "";
							while ((sItem = Console.ReadLine()) != null)
							{ // add to powershell pipeline
								colInput.Add(sItem);
							}
						}
"@
		})
						colInput.Complete();

						PSDataCollection<PSObject> colOutput = new PSDataCollection<PSObject>();
						colOutput.DataAdded += new EventHandler<DataAddedEventArgs>(delegate(object sender, DataAddedEventArgs e)
						{
							ui.WriteLine(colOutput[e.Index].ToString());
						});

						int separator = 0;
						int idx = 0;
						foreach (string s in args)
						{
							if (string.Compare(s, "-wait", true) == 0)
								paramWait = true;
							else if (s.StartsWith("-extract", StringComparison.InvariantCultureIgnoreCase))
							{
								string[] s1 = s.Split(new string[] { ":" }, 2, StringSplitOptions.RemoveEmptyEntries);
								if (s1.Length != 2)
								{
$(if (!$noConsole) { @"
									Console.WriteLine("If you specify the -extract option you need to add a file for extraction in this way\r\n   -extract:\"<filename>\"");
"@
		}
		else { @"
									MessageBox.Show("If you specify the -extract option you need to add a file for extraction in this way\r\n   -extract:\"<filename>\"", System.AppDomain.CurrentDomain.FriendlyName, MessageBoxButtons.OK, MessageBoxIcon.Error);
"@
		})
									return 1;
								}
								extractFN = s1[1].Trim(new char[] { '\"' });
							}
							else if (string.Compare(s, "-end", true) == 0)
							{
								separator = idx + 1;
								break;
							}
							else if (string.Compare(s, "-debug", true) == 0)
							{
								System.Diagnostics.Debugger.Launch();
								break;
							}
							idx++;
						}

						string script = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(@"$($script)"));

						if (!string.IsNullOrEmpty(extractFN))
						{
							System.IO.File.WriteAllText(extractFN, script);
							return 0;
						}

						powershell.AddScript(script);

						// parse parameters
						string argbuffer = null;
						// regex for named parameters
						System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"^-([^: ]+)[ :]?([^:]*)$");

						for (int i = separator; i < args.Length; i++)
						{
							System.Text.RegularExpressions.Match match = regex.Match(args[i]);
							if (match.Success && match.Groups.Count == 3)
							{ // parameter in powershell style, means named parameter found
								if (argbuffer != null) // already a named parameter in buffer, then flush it
									powershell.AddParameter(argbuffer);

								if (match.Groups[2].Value.Trim() == "")
								{ // store named parameter in buffer
									argbuffer = match.Groups[1].Value;
								}
								else
									// caution: when called in powershell $TRUE gets converted, when called in cmd.exe not
									if ((match.Groups[2].Value == "$TRUE") || (match.Groups[2].Value.ToUpper() == "\x24TRUE"))
									{ // switch found
										powershell.AddParameter(match.Groups[1].Value, true);
										argbuffer = null;
									}
									else
										// caution: when called in powershell $FALSE gets converted, when called in cmd.exe not
										if ((match.Groups[2].Value == "$FALSE") || (match.Groups[2].Value.ToUpper() == "\x24"+"FALSE"))
										{ // switch found
											powershell.AddParameter(match.Groups[1].Value, false);
											argbuffer = null;
										}
										else
										{ // named parameter with value found
											powershell.AddParameter(match.Groups[1].Value, match.Groups[2].Value);
											argbuffer = null;
										}
							}
							else
							{ // unnamed parameter found
								if (argbuffer != null)
								{ // already a named parameter in buffer, so this is the value
									powershell.AddParameter(argbuffer, args[i]);
									argbuffer = null;
								}
								else
								{ // position parameter found
									powershell.AddArgument(args[i]);
								}
							}
						}

						if (argbuffer != null) powershell.AddParameter(argbuffer); // flush parameter buffer...

						// convert output to strings
						powershell.AddCommand("out-string");
						// with a single string per line
						powershell.AddParameter("stream");

						powershell.BeginInvoke<string, PSObject>(colInput, colOutput, null, new AsyncCallback(delegate(IAsyncResult ar)
						{
							if (ar.IsCompleted)
								mre.Set();
						}), null);

						while (!me.ShouldExit && !mre.WaitOne(100))
						{ };

						powershell.Stop();

						if (powershell.InvocationStateInfo.State == PSInvocationState.Failed)
							ui.WriteErrorLine(powershell.InvocationStateInfo.Reason.Message);
					}

					myRunSpace.Close();
				}
			}
			catch (Exception ex)
			{
$(if (!$noConsole) { @"
				Console.Write("An exception occured: ");
				Console.WriteLine(ex.Message);
"@
		}
		else { @"
				MessageBox.Show("An exception occured: " + ex.Message, System.AppDomain.CurrentDomain.FriendlyName, MessageBoxButtons.OK, MessageBoxIcon.Error);
"@
		})
			}

			if (paramWait)
			{
$(if (!$noConsole) { @"
				Console.WriteLine("Hit any key to exit...");
				Console.ReadKey();
"@
		}
		else { @"
				MessageBox.Show("Click OK to exit...", System.AppDomain.CurrentDomain.FriendlyName);
"@
		})
			}
			return me.ExitCode;
		}

		static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
		{
			throw new Exception("Unhandled exception in PS2EXE");
		}
	}
}
"@
	#endregion
	
	$configFileForEXE2 = "<?xml version=""1.0"" encoding=""utf-8"" ?>`r`n<configuration><startup><supportedRuntime version=""v2.0.50727""/></startup></configuration>"
	$configFileForEXE3 = "<?xml version=""1.0"" encoding=""utf-8"" ?>`r`n<configuration><startup><supportedRuntime version=""v4.0"" sku="".NETFramework,Version=v4.0"" /></startup></configuration>"
	
	Write-Host "Compiling file... " -NoNewline
	$cr = $cop.CompileAssemblyFromSource($cp, $programFrame)
	if ($cr.Errors.Count -gt 0)
	{
		Write-Host ""
		Write-Host ""
		if (Test-Path $outputFile)
		{
			Remove-Item $outputFile -Verbose:$FALSE
		}
		Write-Host -ForegroundColor red "Could not create the PowerShell .exe file because of compilation errors. Use -verbose parameter to see details."
		$cr.Errors | % { Write-Verbose $_ -Verbose:$verbose }
	}
	else
	{
		Write-Host ""
		Write-Host ""
		if (Test-Path $outputFile)
		{
			Write-Host "Output file " -NoNewline
			Write-Host $outputFile -NoNewline
			Write-Host " written"
			
			if ($debug)
			{
				$cr.TempFiles | ? { $_ -ilike "*.cs" } | select -first 1 | % {
					$dstSrc = ([System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($outputFile), [System.IO.Path]::GetFileNameWithoutExtension($outputFile) + ".cs"))
					Write-Host "Source file name for debug copied: $($dstSrc)"
					Copy-Item -Path $_ -Destination $dstSrc -Force
				}
				$cr.TempFiles | Remove-Item -Verbose:$FALSE -Force -ErrorAction SilentlyContinue
			}
			if (!$noConfigfile)
			{
				if ($runtime20)
				{
					$configFileForEXE2 | Set-Content ($outputFile + ".config") -Encoding UTF8
					Write-Host "Config file for EXE created."
				}
				if ($runtime40)
				{
					$configFileForEXE3 | Set-Content ($outputFile + ".config") -Encoding UTF8
					Write-Host "Config file for EXE created."
				}
			}
		}
		else
		{
			Write-Host "Output file " -NoNewline -ForegroundColor Red
			Write-Host $outputFile -ForegroundColor Red -NoNewline
			Write-Host " not written" -ForegroundColor Red
		}
	}
	
	if ($requireAdmin)
	{
		if (Test-Path $($outputFile + ".win32manifest"))
		{
			Remove-Item $($outputFile + ".win32manifest") -Verbose:$FALSE
		}
	}
}

#Grabs MSI details
#Source: http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/
function getMSIData
{
	param (
		[parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.IO.FileInfo]$Path,
		[parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("ProductCode", "ProductVersion", "ProductName", "Manufacturer", "ProductLanguage", "FullVersion")]
		[string]$Property
	)
	Process
	{
		try
		{
			# Read property from MSI database
			$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
			$MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path.FullName, 0))
			$Query = "SELECT Value FROM Property WHERE Property = '$($Property)'"
			$View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
			$View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
			$Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
			$Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
			
			# Commit database and close view
			$MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
			$View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)
			$MSIDatabase = $null
			$View = $null
			
			# Return the value
			return $Value
		}
		catch
		{
			Write-Host -ForegroundColor Red "Uh oh... your selected file broke me :("; break
		}
	}
	End
	{
		# Run garbage collection and release ComObject
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
		[System.GC]::Collect()
	}
}

#Generates install/uninstall scripts based on 
function generateScript
{
	param (
		[Parameter(Mandatory = $true)]
		$switches,
		[Parameter(Mandatory = $true)]
		$installationFileLocation,
		[Parameter(Mandatory = $false)]
		$productCode,
		[Parameter(Mandatory = $false)]
		$desktopIconName,
		[Parameter(Mandatory = $false)]
		$startMenuShortcut,
		[Parameter(Mandatory = $false)]
		$copyOverFile,
		[Parameter(Mandatory = $false)]
		$copyIntoDirectory,
		[Parameter(Mandatory = $false)]
		$extraCode
	)
	
	#If generation is complete
	$complete = "false"
	
	#Get the current date
	$currentDate = Get-Date -DisplayHint Date
	
	#Set the file name without the extension
	$installationFileName = [IO.Path]::GetFileNameWithoutExtension($installationFileLocation)
	
	#Set the file name with the extension
	$installationFile = Split-Path $installationFileLocation -Leaf
	
    #Set the parent directory of the installation file
    $parentInstallationFile = Split-Path $installationFileLocation -Parent

	if ($installationFileLocation -match '.msi' -and $complete -match 'false')
	{
		
		Write-Host -ForegroundColor Green "Generating Scripts..."
		#Create installation script
		New-Item -Path $parentInstallationFile -Name $("install_" + $installationFileName + ".ps1") -ItemType File -Value $("#Silently installs " + $installationFileName + "`n#Script Auto-Generated by " + $env:USERNAME + ", UWRF" +
			"`n#Date Created " + $currentDate + "
        `n#Silently installs " + $installationFileName + "`nStart-Process" + ' "$PSScriptRoot\' + $installationFile + '"' + " -Wait -ArgumentList " + '"' + $switches + '"') -Force | Out-Null
		
		if ($productCode -ne "")
		{
			#Create uninstallation script
			New-Item -Path $parentInstallationFile -Name $("uninstall_" + $installationFileName + ".ps1") -ItemType File -Value $("#Silently uninstalls " + $installationFileName +
				"`n#Script Auto-Generated by " + $env:USERNAME + ", UWRF" +
				"`n#Date Created " + $currentDate + "
            `n#Silently uninstalls " + $installationFileName + "`nStart-Process " + '"msiexec.exe"' + " -Wait -ArgumentList " + '"/qn /x' + $productCode + '"') -Force | Out-Null
		}
		
		#Create temporary test directory
		if (!(Test-Path $($parentInstallationFile + "\_TEST_YOUR_APPLICATION")))
		{
			New-Item -ItemType Directory $($parentInstallationFile + "\_TEST_YOUR_APPLICATION") -Force
		}
		
		#Create Test Script
		New-Item -Path $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\") -Name $("test_" + $installationFileName + ".ps1") -ItemType File -Value $("#Silently installs " + $installationFileName + "`n#Script Auto-Generated by " + $env:USERNAME + ", UWRF" +
			"`n#Date Created " + $currentDate + "`n`n" + 'Write-Host -ForegroundColor Green "Starting your installation..."' + "`n" + "
        `n#Silently installs " + $installationFileName + "`n" + '$parent = Get-Location | Split-Path -Parent' + "`nStart-Process" + ' "$parent\' + $installationFile + '"' + " -Wait -ArgumentList " + '"' + $switches + '"') -Force | Out-Null
		
		#Copy shortcut onto Desktop
		if ($desktopIconName -ne "UserInputNull")
		{
			
			#Add to install script
			Add-Content $($parentInstallationFile + "\install_" + $installationFileName + ".ps1") -Value $("`n`n#Copies a shortcut onto the public desktop" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $desktopIconName + '" "' + '$env:PUBLIC\Desktop' + '" -Force') -Force | Out-Null
			
			#Add to test script
			Add-Content $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") -Value $("`n`n#Copies a shortcut onto the public desktop" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $desktopIconName + '" "' + '$env:PUBLIC\Desktop' + '" -Force') -Force | Out-Null
		}
		
		#Copy file/folder into directory
		if ($copyOverFile -ne "UserInputNull" -and $copyIntoDirectory -ne "UserInputNull")
		{
			
			#Add to install script
			Add-Content $($parentInstallationFile + "\install_" + $installationFileName + ".ps1") -Value $("`n#Copies a file/folder into a folder" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $copyOverFile + '" ' + '"' + $copyIntoDirectory + '" -Force') -Force | Out-Null
			
			#Add to test script
			Add-Content $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") -Value $("`n#Copies a file/folder into a folder" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $copyOverFile + '" ' + '"' + $copyIntoDirectory + '" -Force') -Force | Out-Null
		}
		
		#Add extra code
		if ($extraCode -ne "UserInputNull")
		{
			
			#Add to install script
			Add-Content $($parentInstallationFile + "\install_" + $installationFileName + ".ps1") -Value $("`n#Added code" + "`n" + $extraCode) -Force | Out-Null
			
			#Add to test script
			Add-Content $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") -Value $("`n#Added code" + "`n" + $extraCode) -Force | Out-Null
		}
		
		#Add to end of test script
		Add-Content $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") -Value $("`n" + 'Write-Host -ForegroundColor Green "Done!"' + "`n" + 'Start-Sleep -s 5')
		
		#Compile Test Scripts
		PS2EXE $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\_TEST_RUN_AS_ADMIN_" + $installationFileName + ".exe") | Out-Null
		
		$complete = "true"
		Write-Host -ForegroundColor Green "`nSuccess!"
	}
	
	if ($installationFileLocation -match '.exe' -and $complete -match 'false')
	{
		
		Write-Host -ForegroundColor Green "Generating Scripts..."
		#Create installation script
		New-Item -Path $parentInstallationFile  -Name $("\install_" + $installationFileName + ".ps1") -ItemType File -Value $("#Silently installs " + $installationFileName + "`n#Script Auto-Generated by " + $env:USERNAME + ", UWRF" +
			"`n#Date Created " + $currentDate + "
        `n#Silently installs " + $installationFileName + "`nStart-Process" + ' "$PSScriptRoot\' + $installationFile + '"' + " -Wait -ArgumentList " + '"' + $switches + '"') -Force | Out-Null
		
		#Create temporary test directory
		if (!(Test-Path $($parentInstallationFile + "\_TEST_YOUR_APPLICATION")))
		{
			New-Item -ItemType Directory $($parentInstallationFile + "\_TEST_YOUR_APPLICATION") -Force
		}
		
		#Create Test Script
		New-Item -Path $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\") -Name $("test_" + $installationFileName + ".ps1") -ItemType File -Value $("#Silently installs " + $installationFileName + "`n#Script Auto-Generated by " + $env:USERNAME + ", UWRF" +
			"`n#Date Created " + $currentDate + "`n`n" + 'Write-Host -ForegroundColor Green "Starting your installation..."' + "`n" + "
        `n#Silently installs " + $installationFileName + "`n" + '$parent = Get-Location | Split-Path -Parent' + "`nStart-Process" + ' "$parent\' + $installationFile + '"' + " -Wait -ArgumentList " + '"' + $switches + '"') -Force | Out-Null
		
		#Copy over shortcut into Start Menu
		if ($startMenuShortcut -ne "UserInputNull")
		{
			
			#Add to install script
			Add-Content $($parentInstallationFile + "\install_" + $installationFileName + ".ps1") -Value $("`n`n#Copies a shortcut into the start menu" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $startMenuShortcut + '" "' + '$env:ProgramData\Microsoft\Windows\Start Menu\Programs' + '" -Force') -Force | Out-Null
			
			#Add to test script
			Add-Content $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") -Value $("`n`n#Copies a shortcut into the start menu" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $startMenuShortcut + '" "' + '$env:ProgramData\Microsoft\Windows\Start Menu\Programs' + '" -Force') -Force | Out-Null
		}
		
		#Copy shortcut onto Desktop
		if ($desktopIconName -ne "UserInputNull")
		{
			
			#Add to install script
			Add-Content $($parentInstallationFile + "\install_" + $installationFileName + ".ps1") -Value $("`n`#Copies a shortcut onto the public desktop" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $desktopIconName + '" "' + '$env:PUBLIC\Desktop' + '" -Force') -Force | Out-Null
			
			#Add to test script
			Add-Content $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") -Value $("`n#Copies a shortcut onto the public desktop" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $desktopIconName + '" "' + '$env:PUBLIC\Desktop' + '" -Force') -Force | Out-Null
		}
		
		#Copy file/folder into directory
		if ($copyOverFile -ne "UserInputNull" -and $copyIntoDirectory -ne "UserInputNull")
		{
			
			#Add to install script
			Add-Content $($parentInstallationFile + "\install_" + $installationFileName + ".ps1") -Value $("`n#Copies a file/folder into a folder" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $copyOverFile + '" ' + '"' + $copyIntoDirectory + '" -Force') -Force | Out-Null
			
			#Add to test script
			Add-Content $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") -Value $("`n#Copies a file/folder into a folder" + "`n" + 'Copy-Item "' + '$PSScriptRoot' + '\' + $copyOverFile + '" ' + '"' + $copyIntoDirectory + '" -Force') -Force | Out-Null
		}
		
		#Add extra code
		if ($extraCode -ne "UserInputNull")
		{
			
			#Add to install script
			Add-Content $($parentInstallationFile + "\install_" + $installationFileName + ".ps1") -Value $("`n#Added code" + "`n" + $extraCode) -Force | Out-Null
			
			#Add to test script
			Add-Content $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") -Value $("`n#Added code" + "`n" + $extraCode) -Force | Out-Null
		}
		
		#Add to end of test script
		Add-Content $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") -Value $("`n" + 'Write-Host -ForegroundColor Green "Done!"' + "`n" + 'Start-Sleep -s 5')
		
		#Compile Test Scripts
		PS2EXE $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\" + "test_" + $installationFileName + ".ps1") $($parentInstallationFile + "\_TEST_YOUR_APPLICATION\_TEST_RUN_AS_ADMIN_" + $installationFileName + ".exe") | Out-Null
		
		$complete = "true"
		Write-Host -ForegroundColor Green "`nSuccess!"
	}
}

#Reference: https://blogs.technet.microsoft.com/heyscriptingguy/2009/09/01/hey-scripting-guy-can-i-open-a-file-dialog-box-with-windows-powershell/
Function Get-FileName($initialDirectory)
{
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
	Out-Null
	
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "Executable Files (*.exe, *.msi)| *.exe;*.msi"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

#XAML Input
$inputXML = @"
<Window x:Name="Script_Generator_Menu" x:Class="ScriptGeneratorGUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:ScriptGeneratorGUI"
        mc:Ignorable="d"
        Title="Script Generator" Height="574.07" Width="585.034" Background="Black" FontFamily="Segoe UI Light" FontSize="16">
    <Grid Margin="0,0,-8,0" Background="#FF851E1E">
        <Grid.RowDefinitions>
            <RowDefinition Height="131*"/>
            <RowDefinition Height="50*"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="IntroText" HorizontalAlignment="Left" Margin="14,10,0,0" TextWrapping="Wrap" Text="Script Generator " VerticalAlignment="Top" Height="22" Width="507" Foreground="White" Grid.Column="3" FontFamily="Segoe UI Black"/>
        <Label x:Name="InstallationFile_Label" Content="Select your installation file:" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top" Foreground="White" Grid.Column="3" Height="31" Width="211" FontFamily="Segoe UI Semibold"/>
        <Label x:Name="Switch_Label" Content="Enter your installation switches:" HorizontalAlignment="Left" Margin="10,70,0,0" VerticalAlignment="Top" Foreground="White" Grid.Column="3" Height="31" Width="247" FontFamily="Segoe UI Semibold"/>
        <TextBox x:Name="InstallationFile_TextBox" Grid.ColumnSpan="4" HorizontalAlignment="Left" Height="27" Margin="285,45,0,0" VerticalAlignment="Top" Width="164" Background="#FFE5E5E5" IsReadOnly="True"/>
        <TextBox x:Name="Switches" Grid.ColumnSpan="4" HorizontalAlignment="Left" Height="27" Margin="285,75,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="250" Background="#FFE5E5E5"/>

        <CheckBox x:Name="StartMenu_CheckBox" Grid.ColumnSpan="4" Content="Add Start Menu shortcut?" HorizontalAlignment="Left" Margin="14,120,0,0" VerticalAlignment="Top" FontWeight="Bold" FontFamily="Segoe UI Semibold" Background="White" Foreground="White"/>
        <CheckBox x:Name="Desktop_CheckBox" Grid.ColumnSpan="4" Content="Add Public Desktop shortcut?" HorizontalAlignment="Left" Margin="12,175,0,0" VerticalAlignment="Top" Foreground="White" FontWeight="Bold" FontFamily="Segoe UI Semibold"/>
        <CheckBox x:Name="Copy_CheckBox" Grid.ColumnSpan="4" Content="Copy a Folder/File?" HorizontalAlignment="Left" Margin="12,230,0,0" VerticalAlignment="Top" Foreground="White" FontWeight="Bold" FontFamily="Segoe UI Semibold"/>
        <CheckBox x:Name="ExtraCode_CheckBox" Grid.ColumnSpan="4" Content="Add extra code?" HorizontalAlignment="Left" Margin="10,309,0,0" VerticalAlignment="Top" Foreground="White" FontWeight="Bold" FontFamily="Segoe UI Semibold"/>

        <Label x:Name="StartMenu_Label" Grid.ColumnSpan="4" Content="Shortcut Filename (including extension):" HorizontalAlignment="Left" Margin="32,140,0,0" VerticalAlignment="Top" Width="253" Foreground="White" FontSize="14" Visibility="Hidden"/>
        <Label x:Name="Desktop_Label" Grid.ColumnSpan="4" Content="Shortcut Filename (including extension):" HorizontalAlignment="Left" Margin="32,195,0,0" VerticalAlignment="Top" Width="253" Foreground="White" FontSize="14" Visibility="Hidden"/>
        <Label x:Name="Copy_Label" Grid.ColumnSpan="4" Content="File/Folder Path:" HorizontalAlignment="Left" Margin="32,250,0,0" VerticalAlignment="Top" Width="253" Foreground="White" FontSize="14" Visibility="Hidden"/>
        <Label x:Name="CopyDestination_Label" Grid.ColumnSpan="4" Content="Destination Path:" HorizontalAlignment="Left" Margin="32,280,0,0" VerticalAlignment="Top" Width="253" Foreground="White" FontSize="14" Visibility="Hidden"/>
        <Label x:Name="ExtraCode_Label" Grid.ColumnSpan="4" Content="Insert other code (use Enter to separate lines):" HorizontalAlignment="Left" Margin="120,329,0,0" VerticalAlignment="Top" Width="329" Foreground="White" FontSize="14" Visibility="Hidden"/>

        <TextBox x:Name="StartMenuShortcut_TextBox" Grid.ColumnSpan="4" HorizontalAlignment="Left" Height="27" Margin="285,143,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="250" Background="#FFE5E5E5" Visibility="Hidden"/>
        <TextBox x:Name="DesktopShortcut_TextBox" Grid.ColumnSpan="4" HorizontalAlignment="Left" Height="27" Margin="285,198,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="250" Background="#FFE5E5E5" Visibility="Hidden"/>
        <TextBox x:Name="Copy_TextBox" Grid.ColumnSpan="4" HorizontalAlignment="Left" Height="27" Margin="185,250,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="250" Background="#FFE5E5E5" Visibility="Hidden"/>
        <TextBox x:Name="CopyDestination_TextBox" Grid.ColumnSpan="4" HorizontalAlignment="Left" Height="27" Margin="185,280,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="250" Background="#FFE5E5E5" Visibility="Hidden"/>
        <TextBox x:Name="ExtraCode_TextBox" Grid.ColumnSpan="4" HorizontalAlignment="Left" Height="118" Margin="32,358,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="503" Background="#FFE5E5E5" AcceptsReturn="True" Visibility="Hidden" Grid.RowSpan="2"/>
        <Label x:Name="Author" Grid.ColumnSpan="4" Content="Austin Webber, UWRF" HorizontalAlignment="Left" Margin="410,9,0,0" VerticalAlignment="Top" Width="125" Foreground="White" FontSize="12"/>
        <Button x:Name="Generate_Button" Content="Generate" Grid.Column="3" HorizontalAlignment="Left" Margin="1,109.009,0,0" VerticalAlignment="Top" Width="568" Height="38.5" Background="White" FontWeight="Bold" IsEnabled="False" Grid.Row="1"/>
        <Border BorderBrush="Black" BorderThickness="3" Grid.ColumnSpan="4" HorizontalAlignment="Left" Height="115" VerticalAlignment="Top" Width="568" Margin="1,0,0,0"/>
        <Border BorderBrush="Black" BorderThickness="3,0,3,3" Grid.ColumnSpan="4" HorizontalAlignment="Left" Height="387" VerticalAlignment="Top" Width="568" Margin="1,110,0,0" Grid.RowSpan="2"/>
        <Button x:Name="Browse_Button" Content="Browse" Grid.ColumnSpan="4" HorizontalAlignment="Left" VerticalAlignment="Top" Height="27" Width="85" Margin="450,45,0,0" BorderBrush="#FFABADB3" Background="#FFE5E5E5"/>
    </Grid>
</Window>

"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try
{
	$Form = [Windows.Markup.XamlReader]::Load($reader)
}
catch
{
	Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
	throw
}


# Load XAML Objects In PowerShell 
$xaml.SelectNodes("//*[@Name]") | %{
	"trying item $($_.Name)";
	try { Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop }
	catch { throw }
} | Out-Null

#Add code to GUI
#If the user doesn't select an installation file
$WPFBrowse_Button.Add_Click({ 
$WPFInstallationFile_TextBox.Text = Get-FileName -initialDirectory $env:USERNAME\Desktop
$WPFGenerate_Button.IsEnabled = $true
})
#$WPFBrowse_Button.Add_Click({ if ($WPFInstallationFile_TextBox.Text -ne $null) { $WPFGenerate_Button.IsEnabled = $true }
#		else { $WPFGenerate_Button.IsEnabled = $false } })

#Handle Start Menu Shortcut CheckBox interactions
$WPFStartMenu_CheckBox.Add_Checked({ if ($WPFStartMenu_Label.Visibility -ne 'Visible') { $WPFStartMenu_Label.Visibility, $WPFStartMenuShortcut_TextBox.Visibility = 'Visible', 'Visible' } })
$WPFStartMenu_CheckBox.Add_UnChecked({ if ($WPFStartMenu_Label.Visibility -eq 'Visible') { $WPFStartMenu_Label.Visibility, $WPFStartMenuShortcut_TextBox.Visibility = 'Hidden', 'Hidden' } })

#Handle Desktop Shortcut CheckBox interactions
$WPFDesktop_CheckBox.Add_Checked({ if ($WPFDesktop_Label.Visibility -ne 'Visible') { $WPFDesktop_Label.Visibility, $WPFDesktopShortcut_TextBox.Visibility = 'Visible', 'Visible' } })
$WPFDesktop_CheckBox.Add_UnChecked({ if ($WPFDesktop_Label.Visibility -eq 'Visible') { $WPFDesktop_Label.Visibility, $WPFDesktopShortcut_TextBox.Visibility = 'Hidden', 'Hidden' } })

#Handle Copy File CheckBox interactions
$WPFCopy_CheckBox.Add_Checked({
		if ($WPFCopy_Label.Visibility -ne 'Visible')
		{
			$WPFCopy_Label.Visibility = 'Visible'
			$WPFCopy_TextBox.Visibility = 'Visible'
			$WPFCopyDestination_Label.Visibility = 'Visible'
			$WPFCopyDestination_TextBox.Visibility = 'Visible'
		}
	})
$WPFCopy_CheckBox.Add_UnChecked({
		if ($WPFCopy_Label.Visibility -eq 'Visible')
		{
			$WPFCopy_Label.Visibility = 'Hidden'
			$WPFCopy_TextBox.Visibility = 'Hidden'
			$WPFCopyDestination_Label.Visibility = 'Hidden'
			$WPFCopyDestination_TextBox.Visibility = 'Hidden'
		}
	})

#Handle Extra Code CheckBox interactions
$WPFExtraCode_CheckBox.Add_Checked({ if ($WPFExtraCode_Label.Visibility -ne 'Visible') { $WPFExtraCode_Label.Visibility, $WPFExtraCode_TextBox.Visibility = 'Visible', 'Visible' } })
$WPFExtraCode_CheckBox.Add_UnChecked({ if ($WPFExtraCode_Label.Visibility -eq 'Visible') { $WPFExtraCode_Label.Visibility, $WPFExtraCode_TextBox.Visibility = 'Hidden', 'Hidden' } })



#Handle Generation Button
$WPFGenerate_Button.Add_Click({
		if ($WPFStartMenu_CheckBox.IsChecked -eq $false) { $WPFStartMenuShortcut_TextBox.Text = "UserInputNull" }
		if ($WPFDesktop_CheckBox.IsChecked -eq $false) { $WPFDesktopShortcut_TextBox.Text = "UserInputNull" }
		if ($WPFCopy_CheckBox.IsChecked -eq $false)
		{
			$WPFCopyDestination_TextBox.Text = "UserInputNull"
			$WPFCopy_TextBox.Text = "UserInputNull"
		}
		if ($WPFExtraCode_CheckBox.IsChecked -eq $false) { $WPFExtraCode_TextBox.Text = "UserInputNull" }
		
		#Gather MSI productCode
		if ($WPFInstallationFile_TextBox.Text -match '.msi') { $productCode = getMSIData -Path $WPFInstallationFile_TextBox.Text -Property ProductCode }
		
		#Generate Script
		generateScript -Switches $WPFSwitches.Text -installationFileLocation $WPFInstallationFile_TextBox.Text -productCode $productCode -desktopIconName $WPFDesktopShortcut_TextBox.Text -startMenuShortcut $WPFStartMenuShortcut_TextBox.Text -copyOverFile $WPFCopy_TextBox.Text -copyIntoDirectory $WPFCopyDestination_TextBox.Text -extraCode $WPFExtraCode_TextBox.Text
		
		#Close form
		$Form.Close()
	})


#Show form
$Form.ShowDialog() | Out-Null
#Developed by Austin Webber, UWRF




