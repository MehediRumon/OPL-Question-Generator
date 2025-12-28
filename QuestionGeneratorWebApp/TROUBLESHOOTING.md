# Troubleshooting Guide - Question Generator Web Application

## Common Issues and Solutions

### Issue: "Could not load file or assembly 'office, Version=15.0.0.0'"

This error occurs when the Microsoft Office Primary Interop Assemblies (PIAs) are not properly installed or registered on the system.

#### Root Cause
The `Microsoft.Office.Interop.Word` package depends on the Office Primary Interop Assembly (`office.dll`), which is NOT included in the NuGet package and must be installed separately on the target system.

#### Solutions (Try in order)

##### Solution 1: Install Office Primary Interop Assemblies
The Office PIAs are typically installed with Microsoft Office, but may need to be installed separately for development/server environments.

**Option A: Install Microsoft Office**
- Install Microsoft Office on the machine where the web application will run
- This is the most reliable solution as it ensures all dependencies are present

**Option B: Install Office PIAs Redistributable (for servers without Office)**
1. Download and install: [Microsoft Office 2013 Primary Interop Assemblies](https://www.microsoft.com/en-us/download/details.aspx?id=40390)
2. Or for Office 2016: [Microsoft Office 2016 Primary Interop Assemblies](https://www.microsoft.com/en-us/download/details.aspx?id=50949)
3. Restart your machine after installation

**Option C: Register Office PIAs in GAC**
If Office is installed but PIAs are not registered:
```powershell
# Run as Administrator
cd "C:\Windows\Microsoft.NET\assembly\GAC_MSIL"
gacutil /i "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16\Microsoft.Office.Interop.Word.dll"
```

##### Solution 2: Rebuild After Configuration Change
After installing PIAs:
```bash
cd QuestionGeneratorWebApp
dotnet clean
dotnet build
dotnet run
```

##### Solution 3: Check App Pool Identity (IIS Deployment)
If deploying to IIS, ensure the Application Pool identity has permissions to access Office:
1. Open IIS Manager
2. Select Application Pool
3. Advanced Settings → Identity
4. Set to a user account that has Office installed and configured
5. Restart Application Pool

##### Solution 4: Enable 32-bit Applications (IIS)
Some Office components require 32-bit mode:
1. Open IIS Manager
2. Select Application Pool
3. Advanced Settings → Enable 32-Bit Applications → True
4. Restart Application Pool

##### Solution 5: Configure DCOM Permissions
Office Interop in server environments requires DCOM configuration:

1. Run `dcomcnfg` (Component Services)
2. Navigate to: Component Services → Computers → My Computer → DCOM Config
3. Find "Microsoft Word 97-2003 Document"
4. Right-click → Properties
5. **Identity Tab**: Select "The interactive user"
6. **Security Tab**: 
   - Launch and Activation Permissions: Add IIS_IUSRS with Local Launch and Local Activation
   - Access Permissions: Add IIS_IUSRS with Local Access
7. Click OK and restart IIS

### Issue: Application runs but Office operations fail

#### Symptoms
- Web app starts successfully
- API endpoints respond
- Error occurs when trying to generate documents

#### Solutions

##### Check Microsoft Word Installation
```powershell
# Verify Word is installed
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe
```

##### Verify Word Can Start
```powershell
# Try starting Word
Start-Process "winword.exe"
```

##### Check File Permissions
Ensure the application has write permissions:
```powershell
# Grant permissions to output folders
icacls "C:\path\to\QuestionGeneratorWebApp\Generated" /grant "IIS_IUSRS:(OI)(CI)F" /T
icacls "C:\path\to\QuestionGeneratorWebApp\Uploads" /grant "IIS_IUSRS:(OI)(CI)F" /T
```

### Issue: COM Exception or "Old format or invalid type library"

#### Solution
This typically occurs when 32-bit/64-bit architecture mismatch exists:

1. Determine Office architecture:
   ```powershell
   # Check if Office is 32-bit or 64-bit
   (Get-ItemProperty HKLM:\Software\Microsoft\Office\ClickToRun\Configuration).Platform
   ```

2. Match application architecture in project file:
   ```xml
   <PropertyGroup>
     <PlatformTarget>x64</PlatformTarget>  <!-- or x86 for 32-bit Office -->
   </PropertyGroup>
   ```

3. Rebuild application

### Issue: "Retrieving the COM class factory failed"

#### Solution
This is a permissions issue with COM components:

1. Create a Desktop folder for the SYSTEM account:
   ```powershell
   # For 64-bit
   mkdir "C:\Windows\System32\config\systemprofile\Desktop"
   
   # For 32-bit
   mkdir "C:\Windows\SysWOW64\config\systemprofile\Desktop"
   ```

2. Grant permissions:
   ```powershell
   icacls "C:\Windows\System32\config\systemprofile\Desktop" /grant "SYSTEM:(OI)(CI)F" /T
   ```

## Development Environment Setup

### Minimum Requirements
- Windows 10/11 or Windows Server 2016+
- .NET 8.0 SDK
- Microsoft Office (Word) installed
- Visual Studio 2022 or VS Code (optional)

### Verification Steps

1. **Verify .NET Installation**
   ```bash
   dotnet --version
   # Should show 8.0.x
   ```

2. **Verify Office Installation**
   ```powershell
   Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*Microsoft Office*"}
   ```

3. **Verify Assembly Registration**
   ```powershell
   # Check if Word Interop is in GAC
   dir "C:\Windows\Microsoft.NET\assembly\GAC_MSIL" | findstr "Microsoft.Office.Interop.Word"
   ```

4. **Test Office Automation**
   Create a test console app:
   ```csharp
   using Microsoft.Office.Interop.Word;
   
   var wordApp = new Application();
   wordApp.Visible = true;
   wordApp.Quit();
   Console.WriteLine("Office automation working!");
   ```

## Server Deployment Checklist

- [ ] Microsoft Office installed on server
- [ ] Office PIAs installed/registered
- [ ] Application rebuilt with `dotnet clean && dotnet build`
- [ ] DCOM permissions configured
- [ ] Application Pool identity configured
- [ ] Desktop folders created for SYSTEM account
- [ ] File permissions granted for Generated/Uploads folders
- [ ] Architecture (32/64-bit) matches Office installation
- [ ] Firewall rules configured (if needed)
- [ ] Test Office automation from server

## Getting Help

If you continue to experience issues after trying these solutions:

1. Check the Windows Event Viewer for detailed error messages:
   - Application logs
   - System logs
   - Security logs (for permission issues)

2. Enable detailed error messages in `appsettings.json`:
   ```json
   {
     "Logging": {
       "LogLevel": {
         "Default": "Debug",
         "Microsoft.AspNetCore": "Debug"
       }
     }
   }
   ```

3. Collect diagnostic information:
   - OS version and architecture
   - Office version and architecture (32-bit/64-bit)
   - .NET version
   - Deployment method (IIS/Kestrel/Azure)
   - Full error stack trace
   - Event Viewer logs

## Known Limitations

- **Windows Only**: Office Interop requires Windows and cannot run on Linux/macOS
- **Office Required**: Microsoft Word must be installed and licensed
- **Performance**: Office automation is resource-intensive; consider request queuing for high-traffic scenarios
- **Concurrent Access**: Office has limited support for concurrent operations; implement proper locking/queuing
- **Server Core**: Windows Server Core editions do not support Office automation

## Alternative Approaches

If Office Interop continues to cause issues, consider these alternatives:

1. **DocumentFormat.OpenXml** - Create/modify documents without Office:
   ```bash
   dotnet add package DocumentFormat.OpenXml
   ```
   - Pros: No Office required, cross-platform, better performance
   - Cons: More complex API, limited to basic operations

2. **Separate Processing Service** - Run Office automation in a dedicated Windows service
   - Pros: Isolated from web app, better error handling
   - Cons: Additional complexity, requires inter-process communication

3. **Azure Functions** - Use Azure Functions on Windows for Office automation
   - Pros: Scalable, managed infrastructure
   - Cons: Additional cost, requires Azure subscription
