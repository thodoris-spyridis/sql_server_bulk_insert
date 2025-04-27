import logging
import pandas as pd
import pyodbc
from app_files import *


def main():

    # Logg any errors in a text file for debugging
    logging.basicConfig(
        filename='error_log.txt',      
        level=logging.ERROR,           
        format='%(asctime)s - %(levelname)s - %(message)s'  
    )

    pd.set_option("future.no_silent_downcasting", True)
    data = pd.read_excel("customer_list.xlsx", sheet_name="Main", dtype=object)
    data = data.fillna("")  # Replace nan values with "" if any

    # In real life senarios a customer list might need cleaning. It is better to clean before inserting.
    # Replace any commas to make sure you don't get errors during the insert into SQL statements. 
    data["CustomerName"] = data["CustomerName"].str.replace("'", "")
    data["Address"] = data["Address"].str.replace("'", "")
    data["LineOffBusiness"] = data["LineOffBusiness"].str.replace("'", "")

    connection_string = "DRIVER={ODBC Driver 17 for SQL Server};SERVER="+SERVER+";DATABASE="+DATABASE+";UID="+USERNAME+";PWD="+PW #+";TrustServerCertificate=yes;"
    # Connect using the connection string
    cnxn = pyodbc.connect(connection_string)
    # Replace the driver with the on you are running. If you are connecting to a cloud SQL server instance, you can just use {SQL Server} for the driver

    cursor = cnxn.cursor()
    cursor.fast_executemany = True # Eneble fast execution to prepare one single batch insert operation in stead of running and commiting one query at a time

    # Queries (? works as a place holder)
    custtable_query = f"""INSERT INTO MainCustTable (
                                CustomerId,
                                CustomerName,
                                RegId,
                                LineOffBusiness)
                            VALUES
                            (?, ?, ?, ?)"""

    contacttable_query = f"""INSERT INTO CustContact (
                                CustomerId,
                                ContactId,
                                Phone,
                                Email)
                            VALUES
                            (?, ?, ?, ?)"""

    addresstable_query = f"""INSERT INTO CustAddress (
                                CustomerId,
                                AddressId,
                                City,
                                Country,
                                Address)
                            VALUES
                            (?, ?, ?, ?, ?)"""

    # loop the excel table and prepare the query params
    custtable_params = []
    contacttable_params = []
    addresstable_params = []

    for row in data.itertuples():
        # Main table params
        custtable_params.append((
            row.CustomerId,
            row.CustomerName,
            row.RegId,
            row.LineOffBusiness
        ))

        # Contact table params
        contact_id = "CON_" + row.CustomerId
        contacttable_params.append((
            row.CustomerId,
            contact_id,
            row.Phone,
            row.dummyemail
        ))

        # Address table params
        address_id = "AD_" + row.CustomerId
        addresstable_params.append((
            row.CustomerId,
            address_id,
            row.City,
            row.Country,
            row.Address
        ))

    # Execute queries
    try:
        cursor.executemany(custtable_query, custtable_params)
    except Exception as e:
        logging.exception("An error occurred")
    try:
        cursor.executemany(contacttable_query, contacttable_params)
    except Exception as e:
        logging.exception("An error occurred")
    try:
        cursor.executemany(addresstable_query, addresstable_params)
    except Exception as e:
        logging.exception("An error occurred")

    # Commit queries and close cursor
    cnxn.commit()
    cursor.close()


if __name__ == "__main__":
    main()