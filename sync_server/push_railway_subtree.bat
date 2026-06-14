@echo off
setlocal

REM Push only the deployable sync server folder to the separate Railway repo.
git subtree push --prefix=sync_server/Hosny-sync-server railway-sync-server main

