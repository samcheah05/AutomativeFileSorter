import mysql.connector
from mysql.connector import Error
from db_connection import create_connection, close_connection

def register_user(username, email, password):
    """
    Registers a new user in the database.
    Returns True if successful, False otherwise.
    """
    conn = create_connection()
    if not conn:
        return False

    try:
        cursor = conn.cursor()
        # Check if the username or email already exists
        cursor.execute("SELECT id FROM users WHERE username = %s OR email = %s", (username, email))
        if cursor.fetchone():
            print("User with that username or email already exists.")
            return False

        # Insert new user record
        cursor.execute(
            "INSERT INTO users (username, email, password) VALUES (%s, %s, %s)",
            (username, email, password)
        )
        conn.commit()
        print("User registered successfully.")
        return True
    except Error as e:
        print(f"Error during registration: {e}")
        return False
    finally:
        cursor.close()
        close_connection(conn)

def login_user(username, password):
    """
    Validates user credentials and returns the user record if login is successful.
    Returns None if credentials are invalid.
    """
    conn = create_connection()
    if not conn:
        return None

    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM users WHERE username = %s", (username,))
        user = cursor.fetchone()
        if user and user['password'] == password:
            print("Login successful!")
            return user
        else:
            print("Invalid credentials.")
            return None
    except Error as e:
        print(f"Error during login: {e}")
        return None
    finally:
        cursor.close()
        close_connection(conn)

# Example usage (for testing):
if __name__ == "__main__":
    # Registration example
    reg_success = register_user("sam", "sam@gmail.com", "sam123")
    
    # Login example (adjust the credentials accordingly)
    user = login_user("sam", "sam123")
    if user:
        print("User details:", user)
