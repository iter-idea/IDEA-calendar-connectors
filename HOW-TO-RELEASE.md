# Release a new version on npm

1. `npm run build` to compile Typescript, check for errors and update the docs.
1. Pull request of the branch with the changes (GitHub).
1. When the pull is confirmed and the code is merged, update the versions in the files (find and replace, e.g. 0.0.1 -> 0.0.2).
1. Commit the files changed in the previous point and commit them using the name of the version
   as commit description (e.g. v0.0.2).
1. Release the new version on npm with `npm run publishPackage`.
