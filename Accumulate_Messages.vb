'D. Ivan Ochoa
'RCET0265
'Fall 2020
'Accumulate_Messages
'https://github.com/ochodieg/Accumulate_Messages


Option Strict On
Option Explicit On
Option Compare Text


Module Accumulate_Messages

    Sub Main()
        Dim message As String

        Dim userInput As String

        Dim clearData As Boolean


        Console.WriteLine($"Enter text to store. 

Enter 'return' at any time to display stored messages.

Enter 'delete' at any time to delete messages")
        Do
            userInput = Console.ReadLine()

            If userInput = "return" Then

                MsgBox(message)

            ElseIf userInput = "delete" Then

                clearData = True

            End If

            message = AccumulateMessage(userInput, clearData)



            clearData = False

        Loop






    End Sub
    Function AccumulateMessage(ByVal newMessage As String, ByVal delete As Boolean) As String
        Static userMessage As String

        If delete Then

            userMessage = ""

        ElseIf newMessage = "return" Then

        Else

            userMessage &= newMessage & vbNewLine


        End If

        Return userMessage


    End Function

End Module
