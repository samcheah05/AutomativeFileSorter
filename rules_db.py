import mysql.connector
from mysql.connector import Error
import json
from db_connection import create_connection, close_connection

def create_rule(user_id, name, source_folder, destination_folder, modifications, is_running=False):
    """
    Insert a new rule into the database.
    'modifications' should be a list, which we'll store as a JSON string.
    Returns the new rule's ID if successful, otherwise None.
    """
    conn = create_connection()
    if not conn:
        return None

    try:
        cursor = conn.cursor()
        modifications_str = json.dumps(modifications)
        query = """
            INSERT INTO rules (user_id, name, source_folder, destination_folder, modifications, is_running)
            VALUES (%s, %s, %s, %s, %s, %s)
        """
        cursor.execute(query, (user_id, name, source_folder, destination_folder, modifications_str, is_running))
        conn.commit()
        rule_id = cursor.lastrowid
        print(f"Rule created successfully with ID: {rule_id}")
        return rule_id
    except Error as e:
        print(f"Error creating rule: {e}")
        return None
    finally:
        cursor.close()
        close_connection(conn)

def get_rules_by_user(user_id):
    """
    Retrieve all rules for a specific user.
    Returns a list of rule dictionaries.
    """
    conn = create_connection()
    if not conn:
        return []

    try:
        cursor = conn.cursor(dictionary=True)
        query = "SELECT * FROM rules WHERE user_id = %s"
        cursor.execute(query, (user_id,))
        rules = cursor.fetchall()
        # Convert JSON string in modifications to a Python list
        for rule in rules:
            rule["modifications"] = json.loads(rule["modifications"]) if rule["modifications"] else []
        return rules
    except Error as e:
        print(f"Error retrieving rules: {e}")
        return []
    finally:
        cursor.close()
        close_connection(conn)

def update_rule(rule_id, **kwargs):
    """
    Update a rule's fields.
    Accepts keyword arguments for fields to update (e.g., name, is_running, files_processed).
    Returns True if update is successful, otherwise False.
    """
    if not kwargs:
        print("No fields provided for update.")
        return False

    conn = create_connection()
    if not conn:
        return False

    try:
        cursor = conn.cursor()
        fields = []
        values = []
        # Special handling for modifications (convert list to JSON)
        if "modifications" in kwargs:
            kwargs["modifications"] = json.dumps(kwargs["modifications"])
        for key, value in kwargs.items():
            fields.append(f"{key} = %s")
            values.append(value)
        values.append(rule_id)
        query = f"UPDATE rules SET {', '.join(fields)} WHERE id = %s"
        cursor.execute(query, tuple(values))
        conn.commit()
        print(f"Rule {rule_id} updated successfully.")
        return True
    except Error as e:
        print(f"Error updating rule: {e}")
        return False
    finally:
        cursor.close()
        close_connection(conn)

def delete_rule(rule_id):
    """
    Delete a rule from the database.
    Returns True if deletion is successful, otherwise False.
    """
    conn = create_connection()
    if not conn:
        return False

    try:
        cursor = conn.cursor()
        query = "DELETE FROM rules WHERE id = %s"
        cursor.execute(query, (rule_id,))
        conn.commit()
        print(f"Rule {rule_id} deleted successfully.")
        return True
    except Error as e:
        print(f"Error deleting rule: {e}")
        return False
    finally:
        cursor.close()
        close_connection(conn)

# Testing the functions (you can run this file directly for testing purposes)
if __name__ == "__main__":
    # Suppose the logged-in user has user_id = 1
    test_user_id = 1

    # Create a new rule
    rule_id = create_rule(
        user_id=test_user_id,
        name="Test Rule",
        source_folder="C:\\Users\\user\\Downloads",
        destination_folder="C:\\Users\\user\\Documents\\test",
        modifications=["date", "file_type"],
        is_running=True
    )

    # Retrieve rules for the user
    rules = get_rules_by_user(test_user_id)
    print("Rules for user:", rules)

    # Update the rule (e.g., change the name and stop it)
    if rule_id:
        update_rule(rule_id, name="Updated Test Rule", is_running=False)

    # Delete the rule
    if rule_id:
        delete_rule(rule_id)
