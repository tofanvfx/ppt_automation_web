import argparse
import sys
import os
import sqlite3

# Add root directory to sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.api.auth import DB_PATH, get_password_hash

def manage_users():
    parser = argparse.ArgumentParser(description="Manage PPT Automation Users")
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # Add user
    add_parser = subparsers.add_parser("add", help="Add a new user")
    add_parser.add_input = add_parser.add_argument("username", help="Username")
    add_parser.add_argument("password", help="Password")
    add_parser.add_argument("--role", choices=["admin", "ppt_generator", "viewer"], default="ppt_generator", help="User role")

    # List users
    subparsers.add_parser("list", help="List all users")

    # Delete user
    del_parser = subparsers.add_parser("delete", help="Delete a user")
    del_parser.add_argument("username", help="Username to delete")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    if args.command == "add":
        try:
            hashed_pw = get_password_hash(args.password)
            cursor.execute(
                "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                (args.username, hashed_pw, args.role)
            )
            conn.commit()
            print(f"✅ User '{args.username}' created successfully with role '{args.role}'.")
        except sqlite3.IntegrityError:
            print(f"❌ Error: Username '{args.username}' already exists.")
            sys.exit(1)

    elif args.command == "list":
        cursor.execute("SELECT id, username, role, created_at FROM users")
        users = cursor.fetchall()
        print(f"\n{'ID':<5} | {'Username':<20} | {'Role':<15} | {'Created At'}")
        print("-" * 65)
        for u in users:
            print(f"{u[0]:<5} | {u[1]:<20} | {u[2]:<15} | {u[3]}")
        print(f"\nTotal users: {len(users)}\n")

    elif args.command == "delete":
        if args.username == "admin":
            print("❌ Error: Cannot delete the default admin user.")
            sys.exit(1)
            
        cursor.execute("DELETE FROM users WHERE username = ?", (args.username,))
        if cursor.rowcount > 0:
            conn.commit()
            print(f"🗑️ User '{args.username}' deleted successfully.")
        else:
            print(f"❌ Error: User '{args.username}' not found.")

    conn.close()

if __name__ == "__main__":
    manage_users()
