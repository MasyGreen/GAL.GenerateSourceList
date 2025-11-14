CHCP 65001
rmdir "build" /s /q
rmdir "dist" /s /q
pyinstaller _scrip.spec
rmdir "build" /s /q