<div align="center">

## ValidateEmail


</div>

### Description

Provides the user with a means of validating an email address. Using the like keyword, checks the email address with a specific pattern and will pick up on invalid chars for example, john.doe@.test.com will return false as would john.doe@com.
 
### More Info
 
strEmail - Email address that needs to be checked

The like statement allows pattern matching against the email address.

Like "*@[a-z,0-9]*.*"

This would allow anything before the @ sign, and the first letter after the @ sign must be either a number of letter and somewhere after that there must be at least one full stop

Boolean value indicating if email address is valid


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Priest](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-priest.md)
**Level**          |Unknown
**User Rating**    |4.3 (166 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-priest-validateemail__1-4181/archive/master.zip)





### Source Code

```
Public Function ValidateEmail(strEmail As String) As Boolean
ValidateEmail = strEmail Like "*@[A-Z,a-z,0-9]*.*"
End Function
```

