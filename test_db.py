from db_connection import create_connection, close_connection

connection = create_connection()

if connection:
    cursor = connection.cursor()
    cursor.execute("SELECT DATABASE();")
    db_record = cursor.fetchone()
    print(f"Connected to Database: {db_record[0]}")

    cursor.close()
    close_connection(connection)
