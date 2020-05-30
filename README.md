# LeaseTransition
Transitional Lease Accounting for IndAS 116 (IFRS 16)

By no means the code is pretty, but it does the job. Also, it does not strictly support variations in the leasing model. 

It takes inputs of rent paid, and the period of lease. Rest is left to the accountant running it. Assuming things like variable lease payments are taken care by the user entering values in input file, the code runs fairly straight forward.

It takes inputs from two files.
Inputs.txt – which tells it about the start of accounting year, as well as the discounting rate to be used.
Inputsheet.xlsx – which tells it about the monthly rental payments which now qualify under the new lease standard. 

It first converts the inputs to a dataframe, adds few columns (discounting rates, present value, interest, lease liability, etc.). Checks if the value of Lease Liability at the end of the lease period reduces to Zero. 

After this is done, the dataframe is split in two more DFs. One for prior periods, and one for Current accounting period.
Prior period is used to generate transitional entries, while current period DF is used to generate / adjust current year entries. It is assumed that the rent paid was booked as expense in P&L.

Based on these three DFs, another DF is created which contains all the journal entries. 

At last, all four DF are exported to a single excel file -which accountants can review/ keep as back up/ etc.
