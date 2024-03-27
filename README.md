
# Our Apps Script

This repository houses custom Google Apps Script projects developed to enhance and automate our workflows within Google Workspace applications.

## License

This project is licensed under the BSD-3-Clause License - see the [LICENSE](LICENSE) file for details.

## Scripts

### Google Docs Default Style Application Script

#### Overview

This script is designed to apply default styles to Google Docs, with a particular focus on normalizing list items and ensuring formatting consistency across documents. It addresses the challenge of maintaining a uniform look and feel in dynamically generated or collaboratively edited documents.

#### Features

-   Resets styles to default for paragraphs and list items.
-   Adjusts list item indents based on nesting level for visual clarity.
-   Applies consistent text formatting across different document sections.

#### Usage

1.  Open a Google Doc.
2.  Navigate to `Extensions > Apps Script`.
3.  Copy and paste the script code into the script editor and save.
4.  Refresh your Google Doc and access the script via `Custom Scripts` in the toolbar.

#### Limitations

Due to Google Docs API constraints, the script encounters several limitations:

-   **Style Application**: The API does not permit fetching or applying the exact default styles programmatically as the Google Docs UI does. Consequently, some styles, such as specific indents and spacing after paragraphs, are approximated based on observed defaults and may not match the UI's "Normal text" setting precisely.
-   **List Item Handling**: Google Docs API provides limited control over list styling, particularly with regard to bullet sizes and list item padding. The script makes best-effort adjustments within these constraints.
-   **Performance**: The script's execution time may vary depending on the document's length and complexity, with longer documents potentially experiencing slower processing times.

### Future Scripts

This section will be updated with additional scripts as they are developed.

## Contributing

We welcome contributions and suggestions to improve our scripts, especially if you found a way to work around the API limitations. Please feel free to fork the repository, make changes, and submit pull requests. For major changes or new script proposals, please open an issue first to discuss what you would like to change.

## Support

These scripts are offered as is.
