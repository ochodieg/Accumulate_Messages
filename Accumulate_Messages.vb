'D. Ivan Ochoa
'RCET0265
'Fall 2020
'Accumulate_Messages
'https://github.com/ochodieg/Accumulate_Messages


Option Strict On
Option Explicit On
Option Compare Text


Module Accumulate_Messages 'Solution, File, Module, Method names all PascalCase - TJR

    Sub Main()
        Dim message As String
        'Remove extra blank lines - TJR
        Dim userInput As String

        Dim clearData As Boolean


        Console.WriteLine($"Enter text to store. 

Enter 'return' at any time to display stored messages.

Enter 'delete' at any time to delete messages")
        Do 'make more general purpose - TJR
            userInput = Console.ReadLine()

            If userInput = "return" Then
                message = AccumulateMessage("", clearData)
                MsgBox(message)

            ElseIf userInput = "delete" Then
                AccumulateMessage(userInput, clearData)
                clearData = True

            Else
                AccumulateMessage(userInput, clearData)
            End If

            'Remove extra blank lines - TJR


            clearData = False

        Loop



        'Remove extra blank lines - TJR


    End Sub
    Function AccumulateMessage(ByVal newMessage As String, ByVal delete As Boolean) As String
        Static userMessage As String

        If delete Then
            userMessage = ""
        ElseIf newMessage <> "" Then 'make more general purpose - TJR
            userMessage &= newMessage & vbNewLine
        End If
        Return userMessage

    End Function

End Module
