
### A small simple script to read data from an excel table and distribute them and insert row in multiple tables in  an SQL Server instance

The real scenario involved a list of possible customers that had to be inserted to a mobile sales app backend. The data was more complex and had to update more tables. The app has a local database that synchronizes and gets data with a data warehouse in a cloud vm. The app dev asked for a certain amount of money, so to cut costs we decided to try to update the new data in the data warehouse with a simple Python script.

To use it first install requirements.txt

```BASH
pip install -r requirements.txt
```

I used pandas to save the excel table data in a data frame. It is best to do some cleaning steps before you insert the data in the database. Real life data can have random characters especially in address fields (it is a customer list in this case) like commas that can break the SQL insert into statements. I did just some replace steps to replace commas for the example.