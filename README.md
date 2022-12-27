# CBS VBA Business Analytics add-in
## Version: 0.0.31
<!-- DO ***NOT*** EDIT ANYTHING ABOVE THIS LINE, INCLUDING THIS COMMENT -->

This repo contains the latest version of the CBS Business Analytics VBA add-in, designed by the [Decision, Risk, and Operations](https://academics.gsb.columbia.edu/phd/academics/dro) at Columbia Business School. See the add-in file itself for copyright and license information.

**Please read this file in detail if you intend to contribute to, or edit, the add-in or its manual**

The workflow to make changes is as follows:
  - Create a new branch for your update. **You should never EVER commit to the main branch directly!**
  - Make the edits you want to make
      - Edit the `.xlam` file (in particular, ensure the version number at the top of `Util.bas` is updated)
      - Edit the user manual LaTeX PDF and compile it (in particular, ensure the version number at the top of the LaTeX file is updated). It is essential that you use unix-style linebreaks (LF rather the CR LF).
      - DO NOT edit the files in the `~VBA Code` folder; these will automatically be updated by the VBA robot
  - Push your branch to GitHub - the VBA robot will get to work and carry out the following actions
      - Ensure the following files exist in the repo
         - `CBS BA Multiplatform add-in.xlam` (the add-in file)
         - `User manual/BA Add-In User Manual.pdf` (the user manual PDF)
         - `User manual/BA Add-In User Manual.tex` (the user manual LaTeX file)
      - Ensure the user manual PDF file was generated from the current version of the user manual LaTeX file (this ensures that you have compiled the LaTeX file before pushing it to the repo)
      - Ensure the version number in `Util.bas` (in the `.xlam` file) matches the version number in the user manual LaTeX file
      - Extract the VBA code from the `.xlam` file into the `~VBA Code` folder
    
    This will take around 2-3 minutes or so; you can track the progress of these steps in the "Actions" tab of github
  - **IMPORTANT**: when the process is done, it will create a new commit in GitHub. *Immediately pull this new commit* so that your local copy isn't behind the origin
  - Look a the `Readme` page for your branch. If any issues were found, it'll contain an `ERROR REPORT` section. If not, you're good to go
  - Once you're happy with your changes, create a pull request to merge your changes into `Main`
  - Finally, when you're ready to release this new version to students, create a new release in GitHub. The name of the release *and* the associated tag should have the form (for eg) `v0.0.31`.
  - A few minutes after creating the release, the user manual PDF and `.xlam` file will be copied into the release on GitHub.