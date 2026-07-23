# Deploying the Table to Shapes add-in (WiX MSI)

The `TableToShapes.Installer` project builds a **per-machine MSI** that installs the assemblies
and writes the COM + PowerPoint `Addins` registry entries (the same registration `install-addin.ps1`
does for a single user, but under `HKLM` for all users). It is currently **unsigned** - see
[Signing](#signing) for what that means and how to add it later.

## Prerequisites

- **WiX Toolset v5** build tooling: `dotnet tool install --global wix` (the `.wixproj` references
  `WixToolset.Sdk/5.0.2`; adjust the version if you have a different one installed).
- The add-in built in the configuration you are packaging (`Release` for shipping).
- On the **target machine**: Windows with **.NET Framework 4.8** and **PowerPoint** installed.
  (The MSI does not install .NET 4.8; wrap it in a WiX **Burn** bundle if you need to chain that.)

## Build

```powershell
# 1. Build the add-in (Release).
dotnet build -c Release

# 2. (Optional) sign the DLLs before packaging - see Signing.

# 3. Build the installer.
dotnet build .\TableToShapes.Installer\TableToShapes.Installer.wixproj -c Release
# -> TableToShapes.Installer\bin\Release\TableToShapes.msi
```

Install / uninstall:

```powershell
msiexec /i TableToShapes.msi           # interactive
msiexec /i TableToShapes.msi /qn       # silent (for managed deployment)
msiexec /x TableToShapes.msi /qn       # uninstall
```

After install, restart PowerPoint and look for **Convert Table** in the **Table to Shapes** group
on the Home tab.

## What the MSI does

- Copies `TableToShapes.AddIn.dll`, `TableToShapes.Interop.dll`, `TableToShapes.Core.dll` to
  `C:\Program Files\TableToShapes`.
- Registers the managed COM class (`mscoree` shim + `CodeBase` pointing at the installed DLL) and
  the `ProgId`, under `HKLM\Software\Classes`.
- Writes `HKLM\...\Office\PowerPoint\Addins\TableToShapes.AddIn.Connect` with `LoadBehavior = 3`.
- Registers for **Add/Remove Programs** with clean upgrade/uninstall (via `MajorUpgrade`).

## Before shipping - things to set / verify

1. **Manufacturer / UpgradeCode.** Set `Manufacturer` in `Package.wxs`. Keep the `UpgradeCode`
   GUID **stable across versions** (it is what lets a new MSI upgrade an old one); bump `Version`
   each release.
2. **Which files ship.** The MSI **harvests every `*.dll`** from the add-in's build output
   (`<Files Include="...\*.dll" />` in `Package.wxs`), so it ships our three libraries **and** the
   Office interop assemblies that the build copies alongside them (`Microsoft.Office.Interop.PowerPoint.dll`,
   `office.dll`, etc.). `.pdb` files are excluded. This is deliberate: `EmbedInteropTypes` is not
   honoured for these particular interop NuGet packages, so the PIA DLLs are needed at runtime and
   must be deployed - harvesting keeps the installer correct regardless of which interop DLLs land
   in `bin`.
3. **`CodeBase` path.** The MSI writes `CodeBase = file:///[installed path]`. This is the one item
   most worth verifying on a real install - confirm the button appears. If fusion rejects the
   path (spaces / backslashes), either install to a space-free folder or switch the value to the
   plain path form.
4. **Bitness.** This is a per-machine MSI writing to the native `HKLM` view. 64-bit Office reads
   it directly; **32-bit Office on 64-bit Windows** resolves COM under `Wow6432Node`, so build a
   32-bit MSI (or add the `Wow6432Node` registry entries) for 32-bit Office estates.
5. **Per-user vs per-machine.** `Scope="perMachine"` installs for all users and needs admin. For a
   no-admin, current-user install, change the scope to `perUser` and the registry roots from
   `HKLM` to `HKCU` (matching `install-addin.ps1`).

## Managed rollout

A silent MSI (`/qn`) deploys cleanly through **Intune**, **SCCM/MECM**, or **Group Policy Software
Installation** - none of which hit SmartScreen/UAC, because they run as SYSTEM. Estates running
**WDAC/AppLocker** with a "signed installers only" rule will require a signed MSI (below).

## Signing

The MSI works unsigned, but there are trust consequences:

- **Unsigned installer:** an interactive user who double-clicks a downloaded MSI/`setup.exe` gets a
  SmartScreen "unknown publisher" warning, and the UAC prompt shows an unknown publisher. Silent
  managed deployment is unaffected.
- **Unsigned add-in DLL:** default Office loads COM add-ins regardless, but a hardened Office with
  a "require add-ins signed by a trusted publisher" policy will install the add-in yet **not load
  it** (visible only in the log as `LoadBehavior` being reset).

To sign (do it in this order so signatures nest correctly):

1. Sign the **three DLLs** after step 1 of the build (see the repo's signing notes; for production
   use an internal-PKI, public-CA hardware-token, or cloud signing certificate - not the
   self-signed dev cert).
2. Build the MSI (WiX harvests the already-signed DLLs).
3. Sign the **MSI**: `signtool sign /fd SHA256 /tr <timestamp-url> /td SHA256 TableToShapes.msi`.
4. If you add a Burn `setup.exe` bundle, sign it last (`insignia` reattaches the engine signature).

Signing is a drop-in: the DLL sign step slots between build and packaging, and the MSI sign is a
final `signtool` call - no changes to the WiX authoring.

## Note on the dev signing script

`TableToShapes.AddIn/sign-dev.ps1` is **separate** from and **not required by** the MSI. It only
matters if you want to test how the add-in behaves under a "signed add-ins only" Office policy on
your dev box. You can ignore it for packaging.
