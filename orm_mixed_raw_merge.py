# pip install python-dotenv 
# pip install SQLAlchemy
# pip install oracledb
# pip install pyodbc
# https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16


import os
from dotenv import load_dotenv
from sqlalchemy import create_engine, select, text, insert, update
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.orm import sessionmaker, declarative_base
from sqlalchemy import Column, String, Numeric, Date

# Load environment variables
load_dotenv()

# Get credentials from environment variables
ORACLE_USERNAME = os.getenv('ORACLE_USERNAME')
ORACLE_PASSWORD = os.getenv('ORACLE_PASSWORD')
ORACLE_SERVER = os.getenv('ORACLE_SERVER')
PLUGGABLE_DB = os.getenv('PLUGGABLE_DB')
MSSQL_SERVER = os.getenv('MSSQL_SERVER')
MSSQL_INSTANCE = os.getenv('MSSQL_INSTANCE')
MSSQL_DB = os.getenv('MSSQL_DB')

# Define separate environment variables for Oracle and MSSQL connections
oracle_connection_string = f'oracle+oracledb://{ORACLE_USERNAME}:{ORACLE_PASSWORD}@{ORACLE_SERVER}:1521/?service_name={PLUGGABLE_DB}'
mssql_connection_string = f'mssql+pyodbc://@{MSSQL_SERVER}\\{MSSQL_INSTANCE}/{MSSQL_DB}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes'

# Oracle database connection
oracle_engine = create_engine(oracle_connection_string, pool_recycle=1800)

# MS SQL Server database connection
mssql_engine = create_engine(mssql_connection_string)

# Define Base for ORM models
Base = declarative_base()

# # ORM classes reflecting the Oracle database schema
# class OracleEmployee(Base):
#     __tablename__ = 'employees'
#     employee_id = Column(Numeric(6,0), primary_key=True)
#     first_name = Column(String(20))
#     last_name = Column(String(25))
#     email = Column(String(25))
#     phone_number = Column(String(20))
#     hire_date = Column(Date)
#     job_id = Column(String(10))
#     salary = Column(Numeric(8, 2))
#     commission_pct = Column(Numeric(2, 2))
#     manager_id = Column(Numeric(6, 0))
#     department_id = Column(Numeric(4, 0))

# class OracleDepartment(Base):
#     __tablename__ = 'departments'
#     department_id = Column(Numeric(4, 0), primary_key=True)
#     department_name = Column(String(30))
#     manager_id = Column(Numeric(6, 0))
#     location_id = Column(Numeric(4, 0))

# class OracleJob(Base):
#     __tablename__ = 'jobs'
#     job_id = Column(String(10), primary_key=True)
#     job_title = Column(String(35))
#     min_salary = Column(Numeric(6,0))
#     max_salary = Column(Numeric(6,0))

# ORM class reflecting the MS SQL Server combined table schema
class MSSQLCombinedEmployee(Base):
    __tablename__ = 'combined_employees'
    employee_id = Column(Numeric(6,0), primary_key=True, autoincrement=False)
    first_name = Column(String(20))
    last_name = Column(String(25))
    email = Column(String(25))
    phone_number = Column(String(20))
    hire_date = Column(Date)
    job_id = Column(String(10))
    salary = Column(Numeric(8, 2))
    commission_pct = Column(Numeric(2, 2))
    department_id = Column(Numeric(4, 0))
    department_name = Column(String(30))
    job_title = Column(String(35))
    min_salary = Column(Numeric(6,0))
    max_salary = Column(Numeric(6,0))

# Create session for Oracle
oracle_session = sessionmaker(bind=oracle_engine)()

# Create session for MS SQL Server
mssql_session = sessionmaker(bind=mssql_engine)()

try:
    # Fetch data from Oracle using raw query wrapped in text()
    oracle_query = text("""
    SELECT e.employee_id, e.first_name, e.last_name, e.email, e.phone_number, e.hire_date, e.job_id, e.salary,
           e.commission_pct, e.department_id, d.department_name, j.job_title, j.min_salary, j.max_salary
    FROM employees e
    INNER JOIN departments d ON e.department_id = d.department_id
    INNER JOIN jobs j ON e.job_id = j.job_id
    """)
    oracle_data = oracle_session.execute(oracle_query).fetchall()

    # Iterate over the fetched data and perform UPSERT on MS SQL Server
    for row in oracle_data:
        # Access row data directly by column names using _mapping
        data = row._mapping

        # Try to find an existing record in MS SQL Server
        stmt = select(MSSQLCombinedEmployee).where(MSSQLCombinedEmployee.employee_id == data['employee_id'])
        existing = mssql_session.execute(stmt).scalar()

        if existing:
            # Update existing record
            update_stmt = (
                update(MSSQLCombinedEmployee)
                .where(MSSQLCombinedEmployee.employee_id == data['employee_id'])
                .values(
                    first_name=data['first_name'],
                    last_name=data['last_name'],
                    email=data['email'],
                    phone_number=data['phone_number'],
                    hire_date=data['hire_date'],
                    job_id=data['job_id'],
                    salary=data['salary'],
                    commission_pct=data['commission_pct'],
                    department_id=data['department_id'],
                    department_name=data['department_name'],
                    job_title=data['job_title'],
                    min_salary=data['min_salary'],
                    max_salary=data['max_salary']
                )
            )
            mssql_session.execute(update_stmt)
        else:
            # Insert new record with explicit column names
            insert_stmt = (
                insert(MSSQLCombinedEmployee)
                .values(
                    employee_id=data['employee_id'],
                    first_name=data['first_name'],
                    last_name=data['last_name'],
                    email=data['email'],
                    phone_number=data['phone_number'],
                    hire_date=data['hire_date'],
                    job_id=data['job_id'],
                    salary=data['salary'],
                    commission_pct=data['commission_pct'],
                    department_id=data['department_id'],
                    department_name=data['department_name'],
                    job_title=data['job_title'],
                    min_salary=data['min_salary'],
                    max_salary=data['max_salary']
                )
            )
            mssql_session.execute(insert_stmt)

    # Commit changes to the MS SQL Server database
    mssql_session.commit()
    print("Merge operation completed successfully.")

except SQLAlchemyError as e:
    print(f"An error occurred: {e}")
    mssql_session.rollback()

finally:
    # Close the sessions
    oracle_session.close()
    mssql_session.close()
    oracle_engine.dispose()
    mssql_engine.dispose()
