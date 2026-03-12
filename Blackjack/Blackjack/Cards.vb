Public Class Cards
    Private Number As String
    Private Colour As String
    Private Type As String
    Public Sub New(col As String, Utype As String, num As String)
        Number = num
        Colour = col
        Type = Utype
    End Sub

    Public Function GetNumber() As String
        Return Number
    End Function
    Public Function ReType() As String
        Return Type
    End Function
    Public Function GetColour() As String
        Return Colour
    End Function
    Public Sub GetCard()

        Console.WriteLine("Card number: " & Number)
        Console.WriteLine("Colour: " & Colour)
        Console.WriteLine("Type: " & Type)
    End Sub
End Class
