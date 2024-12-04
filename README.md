# preppy
Produces unit/lesson plans from an XLSX file

## Description
This is a .NET 8.0 console project published in a public GitHub repository.
The unit/lesson plans are created from an XLSX file in a specific format.
The default layout/style is based on the James Cook University Masters of Teaching and Learning course content.

## Description
Run the EXE file either from the command-line or in Explorer.
Optionally pass in source and destination files.
If no source is provided, the application will look for a single XLSX in the current directory.
If no destination is provided, the application will default to 'output.docx'.

## Updating the application
To checkout code:

    cd c:\snjw\code
    git clone https://github.com/simonnjwalker/preppy.git
    cd preppy

To update GitHub, first put notes into CHANGELOG.md, then:

    cd c:\snjw\code\preppy
    git add .
    git commit -m "1.x.x Code message here"
    git push


