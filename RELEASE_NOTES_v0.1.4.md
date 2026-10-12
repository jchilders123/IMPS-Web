# IMPS Web v0.1.4 — Hospital phone no-wrap correction

Based on the user's v0.1.3 corrected-label web package.

- Hospital phone cells now contain an explicit DOCX `w:noWrap` property.
- Phone numbers use Arial 7 pt with slightly condensed character spacing and nonbreaking separator characters to help the full number fit in its existing column.
- The fix also applies to ground and air ambulance phone cells.
- Preserves the working Google autocomplete country removal and the remaining web features.
- Updates the app display version and service-worker cache to 0.1.4.

Upload the contents of this ZIP to the existing GitHub Pages website root and commit. Generate a **new** ICS-206 Word document to check the change.

The code and DOCX formatting can be inspected, but visual rendering in Word and browser preview should be confirmed on a user's installation.
