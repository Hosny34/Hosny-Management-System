@echo off
setlocal

REM Pull the separate Railway repo back into the local subtree if needed.
git subtree pull --prefix=sync_server/Hosny-sync-server railway-sync-server main --squash

