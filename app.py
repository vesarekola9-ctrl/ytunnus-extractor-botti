      - name: Pick entrypoint automatically
        id: pick
        shell: pwsh
        run: |
          $ErrorActionPreference = "Stop"
          $preferred = @("app.py","ytunnus_dragdrop_bot.py","protestibotti.py")
          $ENTRYPOINT = $null
          foreach ($p in $preferred) { if (Test-Path $p) { $ENTRYPOINT = $p; break } }
          if (-not $ENTRYPOINT) {
            $candidates = Get-ChildItem -File -Filter *.py | Select-Object -First 1
            if ($candidates) { $ENTRYPOINT = $candidates.Name }
          }
          if (-not $ENTRYPOINT) { Write-Host "ERROR: No entrypoint"; exit 1 }
          "ENTRYPOINT=$ENTRYPOINT" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8

      - name: Build EXE with PyInstaller
        shell: pwsh
        run: |
          $ErrorActionPreference = "Stop"
          $ENTRYPOINT = "${{ steps.pick.outputs.ENTRYPOINT }}"
          pyinstaller --noconfirm --clean --onefile --windowed --name "LeadForgeFI" $ENTRYPOINT
