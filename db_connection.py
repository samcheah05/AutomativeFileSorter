import mysql.connector
from mysql.connector import Error
from dotenv import load_dotenv
load_dotenv()  # Load .env file

# Function to connect to your MySQL Database
def create_connection():
    connection = None
    try:
        connection = mysql.connector.connect(
            host=os.getenv('DB_HOST'),
            database=os.getenv('DB_NAME'),
            user=os.getenv('DB_USER'),
            password=os.getenv('DB_PASSWORD')            
        )
        if connection.is_connected():
            print("Successfully connected to MySQL database 'fyp'.")
            return connection
    except Error as e:
        print(f"Error connecting to MySQL database: {e}")
        return None

# Close connection safely
def close_connection(connection):
    if connection and connection.is_connected():
        connection.close()
        print("MySQL connection closed.")
