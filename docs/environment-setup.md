# Environment Setup

This framework can be used by simply including the script within the VPX script editor. However, we recommend a modern workflow for developing tables that uses Git source control and script separation. This approach makes development faster, less error-prone, and more collaborative.

## Tools You Will Need

- **Git** (for version control)
- **VS Code** (for editing code)
- **Node.js** (for running build scripts)

*Note: This guide does not cover installing these tools, as there are many excellent guides available online for each one.*

## VPX Tool

You will also need **vpxtool** by francisdb ([GitHub link](https://github.com/francisdb/vpxtool)). At the time of writing, this workflow uses version 0.13.0 ([download here](https://github.com/francisdb/vpxtool/releases/tag/v0.13.0)).

Download and install vpxtool, then make sure it is available on your command line by adding it to your system's PATH environment variable.

## Cloning the Example Table

1. **Clone the example table repository:**
   [https://github.com/mpcarr/vpx-example-glf](https://github.com/mpcarr/vpx-example-glf)

2. **Open the project in VS Code.**

3. **Open a terminal in VS Code.**

4. **Navigate to the scripts directory:**
   ```
   cd scripts
   ```

5. **Install the Node.js dependencies:**
   ```
   npm install
   ```
6. **Rename the project and reset git history:**
   ```
   npm run rename-project -- NEW_NAME --git
   ```
   This uses script to rename the example table to your new project name. The optional --git command will reset the git commit history so you have a clean repo for your new project.

7. **Build the VPX file:**
   ```
   npm run assemble-vpx
   ```
   This uses vpxtool to build the VPX file from the repository. You sould now have a `MyProjectName.vpx` file in your local directory.

8. **Start the script watcher:**
   ```
   npm run script-watcher
   ```
   This will start a file watcher process that automatically rebuilds the table script used by VPX whenever you make changes. Any edits you make to the scripts will now be automatically updated. To cancel the watch you can press Ctrl+C within the terminal window.

9. **Open the VPX table and press Play!**

You are now ready to develop and test your VPX GLF table with a modern, efficient workflow.