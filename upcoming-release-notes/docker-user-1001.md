---
category: Enhancements
authors: [tabedzki]
---

[BREAKING CHANGE] The Docker server image now starts as a non-root user by default. If you are updating from an older image, start it once with `user: root` so it can take over your existing data folder, or give the folder to uid 1001 yourself. See the [Docker install docs](https://actualbudget.org/docs/install/docker#updating-from-an-older-image).
