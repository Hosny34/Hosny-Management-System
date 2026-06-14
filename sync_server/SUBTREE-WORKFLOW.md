# Sync Server Subtree Workflow

This repo keeps the deployable server code in:

`sync_server/Hosny-sync-server`

That folder is mirrored to the separate Railway deployment repo:

`https://github.com/Hosny34/Hosny-sync-server`

## Remote

The main repo uses this extra git remote:

`railway-sync-server`

## Recommended workflow

Work normally in the main repo.

When you want to update the Railway repo, push only the server folder:

```bat
sync_server\push_railway_subtree.bat
```

Equivalent raw command:

```bash
git subtree push --prefix=sync_server/Hosny-sync-server railway-sync-server main
```

## If the Railway repo ever changes separately

Pull those changes back into the local subtree with:

```bat
sync_server\pull_railway_subtree.bat
```

Equivalent raw command:

```bash
git subtree pull --prefix=sync_server/Hosny-sync-server railway-sync-server main --squash
```

## Notes

- Keep all deployable server files inside `sync_server/Hosny-sync-server`.
- Do not include the local SQLite runtime files in subtree sync decisions.
- `git subtree` is recommended here over submodules because this server folder is part of your normal development flow and should live naturally inside the main repo.
