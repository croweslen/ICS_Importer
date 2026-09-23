## Installation

1. Go to [script.google.com](https://script.google.com) and click **New project**.
2. Rename the project (e.g. `ICS Importer`) by clicking **Untitled project** at the top.
3. In `Code.gs`, select everything (**Ctrl+A**) and delete it, including the default `function myFunction() {}`.
4. Paste in the full contents of `ics-to-calendar.gs`. Line 1 should start with `/**`.
5. Edit the `CONFIG` block at the top if needed (see [Configuration](#configuration)).
6. Save with **Ctrl+S**.
7. In the toolbar's function dropdown, select **`setup`** and click **Run**.
8. Approve the permissions when prompted. Google will warn that the app is unverified because it's your own script: click **Advanced → Go to ICS Importer (unsafe)** → **Allow**.
