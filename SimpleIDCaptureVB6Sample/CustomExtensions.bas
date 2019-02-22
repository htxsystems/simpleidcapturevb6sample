Attribute VB_Name = "Module1"
Imports

Public Module MyExtensions
    ''This procedure gets the <Description> attribute of an enum constant, if any.
    ''Otherwise it gets the string name of the enum member.
    <Extension()> _
    Public Function Description(ByVal EnumConstant As [Enum]) As String
        Dim fi As Reflection.FieldInfo = EnumConstant.GetType().GetField(EnumConstant.ToString())
        Dim aattr() As DescriptionAttribute = DirectCast(fi.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
        If aattr.Length > 0 Then
            Return aattr(0).Description
        Else
            Return EnumConstant.ToString()
        End If
    End Function
End Module

