


Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnClickMe.Click

        Try
            Dim FinalString As String = ""

            'Load strings, Using hard code for now.
            Dim Teststring1 As String = "all is well"
            Dim Teststring2 As String = "ell that en"
            Dim Teststring3 As String = "hat end"
            Dim Teststring5 As String = "end"
            Dim Teststring4 As String = "t ends well"

            'Put the list into a string (realistically should be using some sort of loop, but for now hard code)
            Dim StringList As New List(Of String)
            StringList.Add(Teststring1)
            StringList.Add(Teststring2)
            StringList.Add(Teststring3)
            StringList.Add(Teststring5)
            StringList.Add(Teststring4)
            StringList.Add(Teststring4)


            'But lets start with the easy stuff first (if a string exists in another, if so, then just use the large one and remove the small one out of the equation)
            Dim ContainerRemovedList As List(Of String) = RemoveDuplicateContainer(StringList)

            FinalString = MergeUniqueStringFragments(ContainerRemovedList)

            If Not String.IsNullOrWhiteSpace(FinalString) Then
                MessageBox.Show($"Combined Fragmented String: {FinalString}")
            Else
                MessageBox.Show("Error has occurred, Combined fragmentstring is empty.")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Function RemoveDuplicateContainer(stringlist As List(Of String)) As List(Of String)
        'This function Removes any dupliate values along with any fragment that exists in another fragment.
        Dim result As List(Of String)

        Try
            result = stringlist.Distinct().ToList()
            'Clone a list to prevent loop from breaking when removing items
            Dim ClonedList As List(Of String) = result.Select(Function(x) x.Clone.ToString()).Distinct().ToList()
            For Each item In ClonedList
                For Each item2 In ClonedList
                    If item2.Contains(item) AndAlso item <> item2 Then
                        'Remove item, and stop checking for its existence in another fragment
                        result.Remove(item)
                        Exit For
                    End If
                Next
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            result = Nothing
        End Try

        Return result
    End Function

    Private Function MergeUniqueStringFragments(Stringlist As List(Of String)) As String

        Dim result As String = ""
        Try
            If Stringlist.Count > 0 Then

                While Stringlist.Count > 1

                    'Setup variable and a clone to work with
                    Dim StringSet1 As String = ""
                    Dim StringSet2 As String = ""
                    Dim Similiarity As Integer = 0
                    Dim ClonedList As List(Of String) = Stringlist.Select(Function(x) x.Clone.ToString()).ToList()

                    For Each Set1 In ClonedList
                        'loop through cloned list twice, 
                        For Each Set2 In ClonedList
                            'Dont bother checking yourself, as 
                            If Set1 <> Set2 Then
                                'Declare/Reset variables for each specific run
                                Dim Counter As Integer = 0
                                Dim ResultFound As Boolean = False

                                'Since im checking against length/ Similarity make sure that my similiarty 
                                Do While Counter <> Set1.Length - 1 AndAlso Counter <> Set2.Length - 1 AndAlso ResultFound = False

                                    Counter += 1
                                    If Set1.Substring(Set1.Length - Counter, Counter) = Set2.Substring(0, Counter) Then
                                        'Found a matching case, so mark success and bail out of the loop
                                        ResultFound = True
                                    End If
                                Loop

                                'If we have result, ONLY log it when the similarity score is higher than existing.
                                If ResultFound AndAlso Counter > Similiarity Then
                                    StringSet1 = Set1
                                    StringSet2 = Set2
                                    Similiarity = Counter
                                End If
                            End If
                        Next
                    Next

                    'At this point, we have 2 sets of string with the highest similarity score, so its time to manipulate the string.
                    If Not String.IsNullOrWhiteSpace(StringSet1) AndAlso Not String.IsNullOrWhiteSpace(StringSet2) Then
                        'Remove similar character from Stringset2, and slapit on at the end of StringSet1
                        Dim UpdatedString As String = StringSet1 & StringSet2.Substring(Similiarity)

                        'Remove BOTH Original String from the list, and add in the new updated string to replace it.
                        Stringlist.Remove(StringSet1)
                        Stringlist.Remove(StringSet2)
                        Stringlist.Add(UpdatedString)
                    End If

                End While

                'By the end of the while loop above, you should only have 1 string in the stringlist
                result = Stringlist(0)

            Else
                'Empty list came out, so split out empty string.
                result = ""
            End If

        Catch ex As Exception
            'Error has occurred, capture the error and return blank string.
            MessageBox.Show(ex.ToString)
            result = ""
        End Try

        Return result

    End Function
End Class
