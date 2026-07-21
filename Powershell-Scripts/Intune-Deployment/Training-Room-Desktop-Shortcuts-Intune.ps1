Start-Transcript -Path "C:\Windows\Temp\ShortcutScript.log" -Force

try {
    # Create icons directory
    Write-Host "Creating icons directory..."
    New-Item -ItemType Directory -Path "C:\Windows\System32\icons" -Force | Out-Null

    # Embed icon as Base64 and save to device
    Write-Host "Decoding and saving icon..."
    $IconBase64 = "AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAAMMOAADDDgAAAAAAAAAAAAAAAAAEAAAATAICAaQEBAS3BQUEtwUFBbcFBQW3BQUFtwUFBbcFBQW3BQUFtwUFBbcFBQW3BQUFtwYGBrcGBga3BgYGtwYGBrcFBQW3BQUFtwUFBbcFBQW3BQUFtwUFBbcFBQW3BQUFtwUFBbcFBQW3BQUFtwICAqQAAABMAAAABAAAAGUqKSPqenhn/5OPfv+em5L/oJ2W/6Cdlv+gnZb/oJ2W/6Cdlv+gnZb/oJ2W/6Kfl/+empP/lJGK/5SQiv+UkIr/lJGK/56ak/+in5f/oJ2W/6Cdlv+gnZb/oJ2W/6Cdlv+gnZb/oJ2W/6Cdlv+hnZb/ioeB/y8uLOoAAABlERAO356ahf/c17r/6OPQ//Xw5f/18OX/9fDl//Xw5f/18OX/9fDl//Xw5f/38ub/0MvC/2NhXP9JSEP/SklE/0pJRP9JSEP/Y2Fc/9DLwv/38ub/9fDl//Xw5f/18OX/9fDl//Xw5f/18OX/9fDl//Xw5f/38uf/sa2l/xMSEt8fHhr+pKCL/7aymv/Ev7L/y8a+/8vGvf/Lxr3/y8a9/8vGvf/Lxr3/y8a9/8jDu/9OTEj/Skg+/4uIdf+LiHb/i4h2/4uHdf9KSD7/TkxI/8jDu//Lxr3/y8a9/8vGvf/Lxr3/y8a9/8vGvf/Lxr3/y8a9/83Iv/+4taz/IyIg/gUFBMsWFRLwMTAp/z49Nv8/Pjf/Pz43/z8+N/8+PTb/Pz43/z8+N/8/Pjf/Pj02/ywrJf85ODD/Pz41/z89Nf8/PTX/Pz41/zk4MP8sKyX/Pj02/z8+N/8/Pjf/Pz43/z8+N/8/Pjf/Pz43/z8+N/8/Pjf/MzMt/xgYF/AGBQXLAAAAEwgIB7GUkYT/z8q4/8O+rf/Cvaz/x8Kx/9DKuf/Ev67/wr2s/8K9rP/Cvaz/w7+u/8O+rf/Cvq3/wr6t/8K+rf/Cvq3/w76t/8O/rv/Cvaz/wr2s/8K9rP/Cvaz/wr2s/8K9rP/Cvaz/w76t/8/KuP+UkYT/CAgHsQAAABMAAAADCQkJq62qov+bmJH/ODo0/zxAOP9haln/p7ua/0pQRf89QTn/PkI6/z5COv8+Qjr/PkI6/z5COv8+Qjr/PkI5/z5COf8+Qjn/PkI5/z5COf8+Qjn/PkI5/z5COf8+Qjn/PkI5/z5COf84OjT/m5iR/62qov8JCQmrAAAAAwAAAAMJCQmrraqi/3d0b/81Vi//Zp9Z/2unXv93uWn/Z6Ba/2ScWP9knFj/ZJxY/2ScWP9knFj/ZJxY/2ScWP9knVj/ZJ1Y/2SdWP9knVj/ZJ1Y/2SdWP9knVj/ZJ1Y/2SdWP9knVj/ZqBZ/zVWL/93dG//raqi/wkJCasAAAADAAAAAwkJCautqqL/d3Rv/0VtPf9ysWX/NlQw/y9IKf8wSyv/MEsr/zBLK/8wSyv/MEsr/zBKKv8vSSr/OFYx/3KxZP9/xHD/fsNv/37Db/9+w2//fsNv/37Db/9+w2//fsNv/37Db/+AxnH/RGs8/3d0b/+tqqL/CQkJqwAAAAMAAAADCQkJq62qov93dG//RGs8/3i6av9Xhk3/U4BJ/1OASf9TgEn/U4BJ/1OASf9TgUn/VIJK/1SCSv9ZiU7/d7hp/33Cbv9+wm//fcFu/33Bbv9+wm//fcFu/33Bbv99wW7/fcFu/3/EcP9Eajv/d3Rv/62qov8JCQmrAAAAAwAAAAMJCQmrraqi/3d0b/9EbDz/ca9k/zZSL/8vRyn/L0gp/y9IKf8vSCn/Lkco/zxcNP93uGn/gMVw/3/EcP99wm7/fsJv/2umXv81UzD/M08u/2ihW/9+wm//fcFu/33Bbv99wW7/f8Rw/0RqO/93dG//raqi/wkJCasAAAADAAAAAwkJCautqqL/d3Rv/0RrPP95u2r/W4xQ/1eHTf9Xh03/V4dN/1eHTf9Xhk3/XpFT/3q8a/99wW7/fcFu/33Bbv9+w2//PV41/xlVYP8bXWr/N1Uw/37Cb/99wW7/fcFu/33Bbv9/xHD/RGo7/3d0b/+tqqL/CQkJqwAAAAMAAAADCQkJq62qov93dG//RGo7/3/FcP9/xHD/f8Rw/3/EcP9/xHD/f8Rw/3/EcP9/xHD/fcFu/33Bbv9+wm7/fsJu/3/Eb/9PeEP/GElO/xlOVv9Ibz7/fsNv/33Bbv99wW7/fcFu/3/EcP9Eajv/d3Rv/62qov8JCQmrAAAAAwAAAAMJCQmrraqi/3d0b/9Eajv/f8Rw/33Bbv99wW7/fcFu/33Bbv99wW7/fsNv/3/EcP99wW7/eLpq/3W0Z/91tGb/ebpq/3a2aP8kOCH/HS8c/3OyZf9+wm//fcFu/33Bbv99wW7/f8Rw/0RqO/93dG//raqi/wkJCasAAAADAAAAAwkJCautqqL/d3Rv/0RqO/9/xHD/fcFu/33Bbv99wW7/fsJu/33Bbv9qpF3/S3RC/zFLK/8jNB//Hywc/yEuHv8lNiH/M04t/yEzHf83VTD/fsNv/37Cbv99wW7/fcFu/33Bbv9/xHD/RGo7/3d0b/+tqqL/CQkJqwAAAAMAAAADCQkJq62qov93dG//RGo7/3/EcP99wW7/fcFu/37Cb/93uGn/P2I4/xgfFv8cGh3/KyYr/z03Pf9OR07/VExU/05HT/9AOkH/Gxgb/w0TDP8/Yjj/d7hp/37Cb/99wW7/fcFu/3/EcP9Eajv/d3Rv/62qov8JCQmrAAAAAwAAAAMJCQmrraqi/3d0b/9Eajv/f8Rw/33Bbv99wW7/fMBt/z1fNv8YFhf/OjU6/0VARf9STFL/ZV1l/2hgaP9oYGj/aGBo/2phav85NDn/JSIl/yQhJP89Xjb/fMBt/33Bbv99wW7/f8Rw/0RqO/93dG//raqi/wkJCasAAAADAAAAAwkJCautqqL/d3Rv/0RqO/9/xHD/fcFu/37Db/9noFv/GB4W/zs3PP9GQUb/SUNJ/2BZYP9dVV3/SkVK/0pFSv9cVVz/aGBo/zgzOP8zLzP/WlJa/x0iG/9noFv/fsNv/33Bbv9/xHD/RGo7/3d0b/+tqqL/CQkJqwAAAAMAAAADCQkJq62qov93dG//RGo7/3/EcP99wW7/f8Rw/1F9R/8cGhz/PUBF/0ZBRv9JREn/QDtA/x0aHf8WFBb/FhQW/xwaHP9CPUL/MCww/zMvM/9oYGj/KSYo/1B9Rv9/xHD/fcFu/3/EcP9Eajv/d3Rv/62qov8JCQmrAAAAAwAAAAMJCQmrraqi/3d0b/9FbDz/gcdx/3/Eb/+Bx3H/THZC/x8dH/9GQUb/NzM3/xwaHP8YFhj/OzY7/1pTWv9aU1r/Ozc7/xkXGf8LCgv/KCUo/2FaYf8uKi7/S3ZC/4HHcf9/xG//gcdx/0VsPP93dG//raqi/wkJCasAAAADAAAAAwkJCautqqL/d3Vv/zJRLP9hllT/X5RT/2KZVv87XDT/FhQW/yAdIP8WFRb/NzI3/1tUW/9nX2f/Z19n/2dfZ/9oYGj/T0lP/w4NDv8RDxH/IyEj/xoYGv87WzT/YplW/1+UU/9hllT/MlEs/3d1b/+tqqL/CQkJqwAAAAMAAAADCQgIqa2pof+inpf/QUI9/0dJQv9HSUL/KCsl/wQGA/8ODA7/Mi4y/1hRWP9nX2f/Z19n/2ZeZv9mXmb/Z19n/19YX/8kIST/Mi4y/1ZQVv8yLjL/DQwN/wQGA/8oKyX/R0lB/0dJQv9BQj3/op6X/62pof8JCAipAAAAAw0MDAAAAACEaGZh/9DLwv/OysH/xsK5/3t4c/8oJyb/LSkt/1VPVf9nXmf/Z19n/2ZeZv9mXmb/Zl5m/2ZeZv9jW2P/Kycr/ykmKf9iWmL/Z19n/2ZeZv9WT1b/LSkt/ygmJv97eHP/xsK5/87Kwf/Qy8L/aGZh/wAAAIQMDAwAAAAAAAAAAB4JCQijIyIg2iUlI+wiISD/Kicq/1BKUP9mXmb/Z19n/2ZeZv9mXmb/Zl5m/2ZeZv9mXmb/ZFxk/zMvM/8iHyL/Xlde/2dfZ/9mXmb/Zl5m/2dfZ/9mXmb/UEpQ/yonKv8iISD/JSUj7CMiINoJCQijAAAAHgAAAAAAAAAAAAAAAAAAAAcAAAAWAAAAjy8rL/9mXmb/Z19n/2ZeZv9mXmb/Zl5m/2ZeZv9mXmb/Zl5m/2dfZ/9HQkf/Hhse/1lSWf9nX2f/Zl5m/2ZeZv9mXmb/Zl5m/2ZeZv9nX2f/Zl5m/y8rL/8AAACPAAAAFgAAAAcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABjHx0f+1VPVf9nX2f/Z19n/2ZeZv9mXmb/Zl5m/2ZeZv9mXmb/Zl5m/1pTWv9ZUln/Z19n/2ZeZv9mXmb/Zl5m/2ZeZv9mXmb/Z19n/2dfZ/9WT1b/Hx0f+wAAAGMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA0AAAB5Dg0O3zIuMv9ZUln/Z19n/2dfZ/9mXmb/Zl5m/2ZeZv9mXmb/Z19n/2dfZ/9mXmb/Zl5m/2ZeZv9mXmb/Z19n/2dfZ/9ZUln/Mi4y/w8ND+AAAAB6AAAADQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAoAQEBjBEQEeY3Mzf/W1Rb/2dfZ/9mXmb/Zl5m/2ZeZv9mXmb/Zl5m/2ZeZv9mXmb/Zl5m/2dfZ/9cVVz/NzM3/xIQEuYBAQGMAAAAKAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAMQIBApgVExXsPDc8/15XXv9oX2j/Zl5m/2ZeZv9mXmb/Zl5m/2hfaP9hWWH/PDg8/xUTFewCAgKZAAAAMQAAAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAADoDAgOlGBcY8UE8Qf9gWWD/aGBo/2hgaP9gWGD/UEpQ+UlESboFBAWYAAAAOwAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAcAAABEBAQErx0aHfQ9OD3/PTg9/xwaHPQIBwiiGxgbFAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAAWAEBAacBAQGnAAAAWAAAAAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAGAAAABwAAAA/AAAA/wAAAP+AAAH/4AAH//gAH//+AH///4H/8="

    $IconBytes = [Convert]::FromBase64String($IconBase64)
    [System.IO.File]::WriteAllBytes("C:\Windows\System32\icons\online-learning.ico", $IconBytes)

    # Verify icon was created
    if (Test-Path "C:\Windows\System32\icons\online-learning.ico") {
        Write-Host "Icon created successfully!"
    } else {
        Write-Host "ERROR: Icon file was not created!"
    }
    # Cleanup old shortcuts from previous versions
    $OldShortcuts = @(
    "C:\Users\Public\Desktop\Contoso SharePoint.lnk",
    "C:\Users\Public\Desktop\Helpdesk.lnk"
     )

foreach ($OldShortcut in $OldShortcuts) {
    if (Test-Path $OldShortcut) {
        Write-Host "Removing old shortcut: $OldShortcut"
        Remove-Item -Path $OldShortcut -Force -ErrorAction SilentlyContinue
    }
}
    # Shortcut 2 - SharePoint
    Write-Host "Creating SharePoint shortcut..."
    $WshShell2 = New-Object -ComObject WScript.Shell
    $Shortcut2 = $WshShell2.CreateShortcut("C:\Users\Public\Desktop\EC Training.lnk")
    $Shortcut2.TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    $Shortcut2.Arguments = "https://contoso.sharepoint.com/sites/ITHelpGuidesandCybersecurityTraining/SitePages/ECTraining.aspx"
    $Shortcut2.IconLocation = "C:\Windows\System32\icons\online-learning.ico, 0"
    $Shortcut2.Save()
    Write-Host "SharePoint shortcut created!"

    # Refresh desktop icons
    Write-Host "Refreshing desktop..."
    $code = @"
[System.Runtime.InteropServices.DllImport("Shell32.dll")]
private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
"@
    Add-Type -MemberDefinition $code -Name WinAPI -Namespace SystemManager
    [SystemManager.WinAPI]::SHChangeNotify(0x8000000, 0x1000, [IntPtr]::Zero, [IntPtr]::Zero)
    Write-Host "Done!"

} catch {
    Write-Host "ERROR: $_"
}

Stop-Transcript