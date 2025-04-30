import mysql.connector
from mysql.connector import Error
from db_connection import create_connection, close_connection

def insert_log(rule_id, action, filename):
    """
    Inserts a log entry into the logs table.
    
    :param rule_id: The rule id associated with the log (can be None)
    :param action: A string describing the action (e.g., "Moved", "Deleted")
    :param filename: The filename processed by the action.
    :return: True if insertion is successful, False otherwise.
    """
    conn = create_connection()
    if not conn:
        return False
    try:
        cursor = conn.cursor()
        query = "INSERT INTO logs (rule_id, action, filename) VALUES (%s, %s, %s)"
        cursor.execute(query, (rule_id, action, filename))
        conn.commit()
        print("Log inserted successfully.")
        return True
    except Error as e:
        print(f"Error inserting log: {e}")
        return False
    finally:
        cursor.close()
        close_connection(conn)

def get_logs(limit=100):
    """
    Retrieves the most recent log entries.
    
    :param limit: Number of log entries to retrieve (default is 100).
    :return: A list of log entry dictionaries.
    """
    conn = create_connection()
    if not conn:
        return []
    try:
        cursor = conn.cursor(dictionary=True)
        query = "SELECT * FROM logs ORDER BY timestamp DESC LIMIT %s"
        cursor.execute(query, (limit,))
        logs = cursor.fetchall()
        return logs
    except Error as e:
        print(f"Error retrieving logs: {e}")
        return []
    finally:
        cursor.close()
        close_connection(conn)

# Test the logging functions if running this file directly
if __name__ == '__main__':
    # Test inserting a log entry
    insert_log(rule_id=1, action="Test Action", filename="testfile.txt")
    
    # Retrieve and print the latest logs
    logs = get_logs()
    for log in logs:
        print(log)
