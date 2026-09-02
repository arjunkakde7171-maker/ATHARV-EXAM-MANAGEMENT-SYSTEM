ATHARVKART V3.6 - SQLITE DATABASE LOCK FIX

FIXES APPLIED:
1. SQLite WAL journal mode enabled
2. SQLite busy timeout set to 30 seconds
3. Database writes use retry commit logic
4. Flask debug reloader disabled
5. Database initialization no longer writes on every request
6. Thread-safe SQLite connection configuration added

IMPORTANT:
Before running:
- Close all old CMD windows running ATHARVKART.
- Close old browser tabs if they are continuously refreshing.
- Extract this version to a NEW folder.
- Run RUN_ATHARVKART_V3_6_SQLITE_FIXED.bat

If an old atharvkart.db remains locked:
1. Stop all Python processes from Task Manager.
2. Restart the application.
3. Do not open the same SQLite database from DB Browser while the app is writing.

Open:
http://127.0.0.1:5000
