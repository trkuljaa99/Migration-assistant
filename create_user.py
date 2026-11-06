"""Create or reset a user. The password is read from stdin, so it never appears in a command line:

    echo <password> | python create_user.py someone@example.com "Full Name" --admin
"""
import sys

import store

args = [a for a in sys.argv[1:] if a != "--admin"]
if not args:
    sys.exit(__doc__)
password = sys.stdin.readline().rstrip("\r\n")
if len(password) < 8:
    sys.exit("The password must have at least 8 characters.")
store.add_user(args[0], password, args[1] if len(args) > 1 else "", "--admin" in sys.argv)
print(f"saved {store.clean_email(args[0])}" + (" (admin)" if "--admin" in sys.argv else ""))
