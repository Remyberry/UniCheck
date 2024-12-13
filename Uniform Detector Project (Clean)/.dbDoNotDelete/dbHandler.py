import sqlite3
import bcrypt

# Connect to SQLite database
conn = sqlite3.connect("unicheck.db")
cursor = conn.cursor()

# Create students table
cursor.execute("""
CREATE TABLE IF NOT EXISTS students (
    student_id TEXT PRIMARY KEY,
    student_name TEXT,
    id_image_path TEXT
)
""")
conn.commit()

# Create operators table
cursor.execute("""
CREATE TABLE IF NOT EXISTS operators (
    operator_id INTEGER PRIMARY KEY AUTOINCREMENT,
    username_or_email TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL
)
""")
conn.commit()


def add_operator(username_or_email, password):
    # Hash the password
    hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())

    # Insert the operator into the database
    conn = sqlite3.connect("unicheck.db")
    cursor = conn.cursor()

    try:
        cursor.execute("""
        INSERT INTO operators (username_or_email, password_hash) VALUES (?, ?)
        """, (username_or_email, hashed_password))
        conn.commit()
        print("Operator added successfully.")
    except sqlite3.IntegrityError as e:
        print("Error: Username or email already exists.")
    finally:
        conn.close()


def add_student(student_id, student_name, image_path):
    # Insert the operator into the database
    conn = sqlite3.connect("unicheck.db")
    cursor = conn.cursor()

    try:
        cursor.execute("""
        INSERT INTO students (student_id, student_name, id_image_path) VALUES (?, ?, ?)
        """, (student_id, student_name, image_path))
        conn.commit()
        print("Data inserted successfully.")
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
    finally:
        conn.close()

# Insert example data
add_operator("admin","adminadmin#4321")
add_operator("admin2","admin22")

add_student("24-0001","Dexter Tiglao", ".dbDoNotDelete\\studentIMGs\\24-0001.jpeg")
add_student("24-0002","Danica Mercado", ".dbDoNotDelete\\studentIMGs\\24-0002.jpeg")
add_student("24-0003","Ralph Pangan", ".dbDoNotDelete\\studentIMGs\\24-0003.jpeg")
