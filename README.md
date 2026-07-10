## Prioritization Tools

### Welcome! 		

This workbook has been built and configured so that it can be easily downloaded, customized, and launched for any given team or workplace environment where there are projects to manage across both people and products. 		
        
### Structure & Guidelines		

Each product has its own worksheet with associated work items (e.g. Product A, Product B, Product C). All new product worksheets should be copied from the Product Template to maintain structure and formatting. 

In addition to each Product sheet, there are also multiple Priority/Staging sheets that can be configured. These lists are downstream from the Product lists and show filtered compilations of work items, which for example could be marked as a Priority or assigned to a specific person. These sheets are governed by power query automations that reference specific flags on work items across all of the Product sheets.

For example, Analyst 1 might be working on 4 projects across 3 different products. If all of their work items are marked with the Flag '1' (or any other symbol), they can all be populated in their own individual list. 	

As many or as few of these lists can be created as needed. They should all follow the same template, but will all need to be configured via Power Query automation. 	

Please note that users should NOT edit the work items in the Priority/Staging lists, as the list of work items within each Product sheet is the real source of truth, whereas Priority/Staging lists are compilations from those lists. 


        
### New Configurations		

Two types of Power Query configurations used here. The first (and simpler) kind opens a connection to an existing table so that it can be used in the second type, which appends multiple tables together via their connections and then filters the list for a specific flag.

    • In order to set up a new connection using Power Query: Data > Get Data > From Other Sources > From Table/Range > PQ Opens Up > Name the Query > Close and Load To > Create Connection Only

    • Once the connections exist, create an Appended query, sourcing from all of the connections (queries) opened, then filter the appended list as needed (based on the Flag field).

    • Once the power query output has been generated, apply the VBA Macro ApplyPriorityTemplate() for best formatting.

    • Note: In order to change column names or order, all power queries will need to be updated in addition to the templates and the pages themselves. 	


In order to add a new Product list, copy the Product template sheet for structure and formatting, open a new connection to the new table, and then update the Power Query automations to append the new source so that new work items created for the Product will appear in the existing compiled Priority/Staging lists.

In order to add a new Priority list, create a new Power Query (or duplicate an existing one) and set the desired flag so that flagged work items in Product lists can appear in the new Priority list as-expected.	

        
### Sheet Formatting		

Each Sheet has a few specific configurations to ensure continuity in structure and formatting.		

    • The Priority column has defined values (via data validation - list) and conditional formatting based on the defined values	

    • The Status column has defined values (via data validation - list) and conditional formatting based on the defined values	

    • The % Complete column is formatted as % and has Data Bars (conditional formatting)

        • Note: A 0% and a 100% may need to be present as anchors so that the color bars work effectively relative to each other. This gets easier with historical items and filtered lists.

    • The Template sheets are purposely defined up to 100 rows. If & when more rows are needed, copy or extend the rows in the table and ensure the formatting is extended for all columns	

        • Note: This may require updating the power query automation if the table doesn't expand automatically.

    • Columns 'Project Description', 'Next Task', & 'Notes' should all have Wrap-Text enabled. However, it is not dynamic, so after making changes (including editing column width), re-Wrapping-Text may be needed.	

While the Worksheets should mostly stay the same across instances, column widths and actual content can be different across sheets. However, Power Query needs the columns to remain in a predictable order, and clean formatting should be used across sheets for ease and efficiency. 		

    • In order to support this, and since Power Query needs to create its own output, the ApplyPriorityTemplate() VBA Macro was created to ensure that it can always be reset to a standard. 	

For best practice, Lock the priority worksheets to ensure that only the source is edited. 		

        
### How to Use		

• Configure the People and Products whose work items will be tracked in the list over time

• Revisit the list as needed to make updates. Refresh the data by going to Data > Refresh All 




