#Project title: sales_order_automation
#Video demo (youtube link): 
#Description: A code to create sales orders in the SAP system automatically, #based on information from an excel file.
#######################
During my time working in the area of production line programming at a company in Brazil, I found that there were many manual processes that required a lot of time to carry out.
My job was to create sales orders within the SAP (System Applications and Products in Data Processing) system, which is nothing more than an RP system that is divided into numerous modules. 
This system was responsible for dictating the production line, but in order to be able to continue with the production process, it was necessary to create sales orders for the products. This activity required many hours of work, because in addition to the payment terms, the types of orders had to be changed according to each type of product purchased by the customer. 
Some time later, when I moved to another area, I was able to develop what #they call an RPA (Robotic Process Automation) to automate this activity, so the tool I have developed here is a reflection of this need, which was observed in my previous job as an SAP user.

At the first moment (1.) the command initializes the SAP system. It's worth remembering that we can create a password authenticator so that the user enters their password and enters the SAP system, or we can instruct the user to open their SAP system before starting the program.

In the second stage (2.), a dataframe was defined which contains the information on the requests to be created. This dataframe can be extracted from excel, csv, etc. In addition, it is worth noting that the information that makes up the sales order will depend directly on the #business rule, as each company may have its own particularities, so if it were to be used for other companies it would be necessary to adapt the fields according to their needs.

The third moment (3.) will start with the logic itself, a for loop has been created for iteration, so if there are more than one record in the spreadsheet the idea is to make it execute all the records and at the end successfully execute all the sales orders.
With each line, it performs a step within the SAP system, simulating the user's actions. The SAP script recorder was used to obtain the SAP steps. 
It will access transaction VA01, which in the SAP ECC version is the transaction used to create sales orders. After accessing it, it will define the type of order, add information in each field relating to the plant, sales area, when the selection is complete, it will enter another screen and fill in the following information: Order type, Order issuer, Goods recipient, Customer reference, Payment code, Incoterms, Material and Order quantity.
At the end, it will save the order and display the generated order on the screen.
It's worth remembering that for this situation I've only placed one order and defined some information as standard, however it can be #adapted to your needs.
##################







