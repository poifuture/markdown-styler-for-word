# MS Word Add-in Markdown Style

This project is created from
[Office-Addin-TaskPane-React](https://github.com/OfficeDev/Office-Addin-TaskPane-React)
template
[commit d23208](https://github.com/OfficeDev/Office-Addin-TaskPane-React/tree/d232082509522d1bf12da371e10676a721ecd247).
We are expecting not to change the default workflow

## TypeScript

[TypeScript](http://www.typescriptlang.org/).
[Office-Addin-TaskPane-React-JS](https://github.com/OfficeDev/Office-Addin-TaskPane-React-JS).

## Debugging

### In chrome (recommend)

1. Open PowerShell !!!
1. yarn install
1. yarn dev-server (start server at localhost:3000)
1. Open chrome, navigate to [Word Online](https://www.office.com/launch/word)
1. Open a new doc -> Insert -> Office Add-ins
1. Upload ./manifest.xml

### Sideload from on-premise MS Word

1. Open PowerShell
1. yarn install
1. yarn start
1. Attach debugger to Internet Explorer with Visual Studio

### More debugging info

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Additional resources

- [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- More Office Add-in samples at
  [OfficeDev on Github](https://github.com/officedev)
