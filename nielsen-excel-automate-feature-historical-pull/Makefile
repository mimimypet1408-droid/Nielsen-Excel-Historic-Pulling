PS_EXEC = pwsh -ExecutionPolicy bypass -File scripts\refreshpath-exec.ps1

install-os-deps:
	winget install --disable-interactivity Schniz.fnm Python.Python.3.13 astral-sh.uv || echo "Packages already installed"
	$(PS_EXEC) fnm install 22
	pwsh -ExecutionPolicy bypass -File scripts\setup-fnm-profile.ps1
	pwsh -ExecutionPolicy bypass -Command "Invoke-WebRequest 'https://github.com/microsoft/WinAppDriver/releases/download/v1.2.1/WindowsApplicationDriver_1.2.1.msi' -OutFile '$(TEMP)\WindowsApplicationDriver_1.2.1.msi'"
	msiexec.exe /i "$(TEMP)\WindowsApplicationDriver_1.2.1.msi" /passive /norestart

install-deps:
	$(PS_EXEC) uv sync
	$(PS_EXEC) fnm-env npm i
	$(PS_EXEC) fnm-env node_modules\.bin\appium driver install windows

bootstrap: install-os-deps install-deps

run-appium:
	node_modules\.bin\appium

.PHONY: install-os-deps install-deps bootstrap run-appium