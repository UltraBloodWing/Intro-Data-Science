TASK 1

Open test.xlsx
Clear out all data in column B
B1: Set to Survived
B2: Set to =OFFSET([titanic_full_survivor_list.xls]titanic3!$B$1, MATCH(D2, TRIM(CLEAN([titanic_full_survivor_list.xls]titanic3!$C$2:$C$2000)), 0), 0)
>>> Explanation: grab the original list, remove spaces before and after
    names and remove non-printable characters, then find the
    passenger index in titanic_full_survivor_list.xls, then
    use OFFSET to get the survival value from titanic_full_survivor_list.xls.
Extend B2 to whole column
Copy column B into column B as values

TASK 2

Open train.xlsx

Rename worksheet "data" to "Original Data"
Copy worksheet "Original Data" into "Discretized"
Insert new column before column G
G1: Set to Age_Group
G2: Set to =IF(ISBLANK(F2), "N/A", IF(F2<14, "Child", IF(F2<35, "Young Adult", IF(F2<50, "Middle-Aged Adult", "Senior"))))
>>> Explanation: G2 becomes N/A if the age is absent, otherwise it displays "Child"
    for those under 14, then "Young Adult" for those aged 14-34, "Middle-Aged Adult" for those aged 35-49,
    and "Senior" for those 50 and above.
Extend G2 to whole column
Insert new column before column I
I1: Set to Has_SibSp
I2: Set to =IF(H2>0, "Has Sibling/Spouse", "No Sibling/Spouse")
Extend I2 to whole column
Insert new column before column K
K1: Set to Has_ParCh
K2: Set to =IF(J2>0, "Has Parent/Child", "No Parent/Child")
Extend K2 to whole column
Insert new column before column N
N1: Set to Fare_Group
N2: Set to =IF(M2<=10, "Low Fare", IF(M2<=20, "Medium Fare", "High Fare"))
Extend N2 to whole column
Q1: Set to Embarked_Refined
Q2: Set to =IF(ISBLANK(P2), "N/A", P2)
Extend Q2 to whole column

Create new worksheet "Probability Tables"
Set A1 to Survived
Set A2 to PivotTable (Table/Range = Discretized!$B$1:$B$892)
Set A2 Rows to Survived, Values to Count of Survived
Set D2 to P(Class)
Set D3 to =B3/B$5
Extend D3 to D4
Set E2 to ln(P(Class))
Set E3 to ln(D3)
Extend E3 to E4

Set A7 to Pclass
Set A8 to PivotTable (Table/Range = Discretized!$A$1:$Q$892)
Set A8 Rows to Pclass, Columns to Survived, Values to Count of Survived
Set F8 to Laplace Smoothing
Set F9 to Survived: 0
Set G9 to Survived: 1
Set F10 to =B10+1
Extend F10 to G10, then to G12
Set E13 to Total
Set F13 to =SUM(F10:F12)
Extend F13 to G13
Set H8 to Conditional Probabilities
Set H9 to P(A|Class=0)
Set I9 to P(A|Class=1)
Set H10 to =F10/F$13
Extend H10 to I10, then to I12
Extend G13 to I13
Set J8 to Natural Log of Probabilities
Set J9 to ln(P(A|Class=0))
Set K9 to ln(P(A|Class=1))
Set J10 to =LN(H10)
Extend J10 to K10, then to K12
<Repeat several times below, except instead of Pclass use Sex, then Age_Group (but hide the N/A row),
Has_SibSp, Has_ParCh, Fare_Group and Embarked_Refined (but hide the N/A row).>

Copy worksheet "Discretized" into "Model 1 Train"
Copy column G into column G as values
<Repeat for columns I, K, N and Q.>
Remove columns "PassengerId", "Age", "SibSp", "Parch", "Ticket", "Fare", "Cabin" and "Embarked"
Swap column B with C

J1: Set to ln(P(Pclass|Class=0))
J2: Set to =IF(C2="N/A", 0, INDEX('Probability Tables'!$J$10:$J$12, MATCH(C2, 'Probability Tables'!$A$10:$A$12, 0)))
>>> Explanation: If C2 is N/A then return 0 to avoid altering the result, otherwise
    get the list of Pclass, then find the row number for the desired Pclass,
    then get the ln(P(A|Class=0)) column for Pclass and retrieve the same row number.
Extend J2 to whole column
<Repeat for the next columns, but use Sex, then Age_Group, Has_SibSp, Has_ParCh, Fare_Group and Embarked.>
Q1: Set to ln(P(Class=0))
Q2: Set to ='Probability Tables'!$E$3
Extend Q2 to whole column
R1: Set to ln(P(Class=0|A))
R2: Set to =SUM(J2:Q2)
Extend R2 to whole column
<Repeat everything after column swap, but with Class=1 instead of Class=0.>

AB1: Set to Normalized % Class=0
AB2: Set to =EXP(R2)/(EXP(R2)+EXP(AA2))*100
Extend AB2 to whole column
AC1: Set to Normalized % Class=1
AC2: Set to =EXP(AA2)/(EXP(AA2)+EXP(R2))*100
Extend AC2 to whole column
AD1: Set to Predicted Class
AD2: Set to =IF(AB2>AC2, 0, 1)
Extend AD2 to whole column
Move column A to AE
Remove column A
AD1: Set to Actual Class

Copy worksheet "Model 1 Train" to "Model 2 Train"
Remove extra columns
<Repeat until all six models are trained.>