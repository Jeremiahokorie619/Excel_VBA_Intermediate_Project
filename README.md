# Excel_VBA_Intermediate_Project
This Project aims to demonstrate the effectiveness of Excel VBA in automating simple, critical, recurring but boring excel data analyst tasks.

### INTRODUCTION

As a data analyst, part of my job includes a lot of monotonous tasks that could be simple but having to do them a considerably large number of times manually could lead to errors. This is the main reason for Excel VBA Macros. VBA stands for "Visual Basic Application" and Excel Macros are simply pockets of codes in Visual Basic language that are used to accomplish a task ranging from simple to extremely complex. 

Literally everything that can be done in excel could also be turned into a macro in order to simply achieve the same result using the macro rather than going through all the various steps every time you have to accomplish that same task.

### OVERVIEW / CASE STUDY
here is a screenshot to aid understanding

![image](https://github.com/user-attachments/assets/d28c45c1-3db9-499d-8b67-29271b01a06f)

The Image above shows the daily Expenses for one retail store like Costco. 

As a Data Analyst at this Retail company. One of your major tasks just before you go home for the weekend is to use the Expenses data for the entire week to create a simple Table like the one below.

![image](https://github.com/user-attachments/assets/a3d85239-fd5b-49b6-9f09-dd4b3ea9f901)

THIS IS WHAT WE ARE TRYING TO CREATE

![image](https://github.com/user-attachments/assets/1eef73ec-24b4-4c17-b03a-873df1d6f36e)


The table gives a sum total of all the expenses for that week **for one store**. Sure, we could Simply Use Copy and Paste From all the previous sheets, plus a little bit of formatting to easily complete this task. But here's the catch, the information in this whole worksheet is **for ONE STORE only**. What if we had to do this same table for 20 or 30 stores before we run off to enjoy our weekend?, Terrible right?. Now This is Where VBA Macros shine!!. With one simple click, we could create the same result for as many stores(with different values) as needed. 

Excel VBA Macros also helps us to eliminate the risk of mistakes as long as our code works properly.

### MACRO CODES
We will need a total of Five(5) Macros to complete this task.

#### 1. "INSERT TEXT" MACRO
This macro creates the framework and determines the amount of space we will be needing to complete this task, 
and it also inserts the headers and **DAYS** of the week.

```vba
Public Sub InsertTxts()
Range("A1").Select
Selection.Value = "DAY"
Selection.Font.Bold = True
Range("B1").Select
Selection.Value = "EXPENSE"
Selection.Font.Bold = True
Range("C1").Select
Selection.Value = "AMOUNT"
Selection.Font.Bold = True
Range("A2").Select
Selection.Value = "Monday"
Selection.Font.Bold = True
Selection.Offset(5, 0).Select
Selection.Value = "Tuesday"
Selection.Font.Bold = True
Selection.Offset(5, 0).Select
Selection.Value = "Wednesday"
Selection.Font.Bold = True
Selection.Offset(5, 0).Select
Selection.Value = "Thursday"
Selection.Font.Bold = True
Selection.Offset(5, 0).Select
Selection.Value = "Friday"
Selection.Font.Bold = True
Selection.Offset(5, 0).Select
Selection.Value = "Total"
Selection.Font.Bold = True
Selection.Offset(5, 0).Select
Selection.Value = "Grand Total"
Selection.Font.Bold = True
End Sub
```
#### 2. "COPY AND PASTE" MACRO
This macro copies data from all the expenses tables and pastes the data strategically on the required sheet

```vba
Public Sub CopyandPaste()

Dim x As Integer

For x = 1 To Worksheets.Count - 2
Worksheets(x).Select
Range("A1").Select
Selection.CurrentRegion.Copy
Sheets("Create my Own").Select
Range("B" & x * 5 - 3).Select
ActiveSheet.Paste
Next x
Range("A1").Select
End Sub
```

#### 3. "TOTALS" MACRO
This macro adds up all the corresponding expenses type- For instance: 
All the rent expenses are added up for the entire week due to this macro

```vba
Public Sub Totals()

Range("C27").Select
Selection.Value = Range("C2").Value + Range("C7").Value + Range("C12").Value + Range("C17").Value + Range("C22").Value
Selection.Font.Bold = "True"
Range("C28").Select
Selection.Value = Range("C3").Value + Range("C8").Value + Range("C13").Value + Range("C18").Value + Range("C23").Value
Selection.Font.Bold = "True"
Range("C29").Select
Selection.Value = Range("C4").Value + Range("C9").Value + Range("C14").Value + Range("C19").Value + Range("C24").Value
Selection.Font.Bold = "True"
Range("C30").Select
Selection.Value = Range("C5").Value + Range("C10").Value + Range("C15").Value + Range("C20").Value + Range("C25").Value
Selection.Font.Bold = "True"
Range("B27").Select
Selection.Value = "Staff"
Selection.Font.Bold = "True"
Range("B28").Select
Selection.Value = "Utilities"
Selection.Font.Bold = "True"
Range("B29").Select
Selection.Value = "Restock"
Selection.Font.Bold = "True"
Range("B30").Select
Selection.Value = "Rent"
Selection.Font.Bold = "True"
End Sub
```

#### 4. GRAND TOTAL MACRO
This macro adds up all the totals for each expense during the week to give a grand total amount representing the sum total of expenses for running that particular store for one week.

```vba
Public Sub GrandTotal()
Range("C32").Select
Selection.Value = Range("C27").Value + Range("C28").Value + Range("C29").Value + Range("C30").Value
Selection.Font.Bold = "True"
End Sub
```

#### 5. "RUN ALL MACROS" MACRO
This is a special macro that runs all the other macros created for this project. This is great because otherwise we would have to run each of the other macros one by one- which reduces the efficiency and negates the main point of the project. Now all we need to do is run this particular macro and it should automatically run the others.

```vba
Public Sub RunAllMacros()
Call InsertTxts
Range("A1").Select
Call CopyandPaste
Range("A1").Select
Call Totals
Range("A1").Select
Call GrandTotal
Range("A1").Select
Columns("A:C").EntireColumn.AutoFit
End Sub
```

### TESTING RESULTS

![image](https://github.com/user-attachments/assets/c4fd03cf-74a5-456f-ad57-6d9160d0f11d)

As soon as the code is run(by clicking the "play" icon highlighted in yellow and circled with red) the table on the left hand side immediately appears like it has been pre_made.

The interesting thing is that the same code could be applied to multiple files and will produce similar results as long as the format is consistent and there are no changes in format of the original files- though the values could be different and it should not bother us at all.

we could save HOURS  of time and a lot of critical errors simply by running the code 20 times than manually creating the table 20 times!!

This is the power of Excel VBA!!

#### Want to test it out for yourself?

1. Simply download the "vba practice backup" file
2. Click on the "Developer" section
3. Click on "Visual Basic"
4. Navigate to the Last Sheet called "Create my Own"
5. Run the Last code named "RunAllMacros()" simply by navigating to it and placing your cursor at the beginning part then clicking on the "Run" icon (The one circled in red in the Testing Results section screenshot) 


