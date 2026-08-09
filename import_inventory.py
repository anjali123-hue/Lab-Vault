"""Import the two uploaded LabVault workbooks into the configured database.

The application imports these files safely on startup as well. This command is
provided for explicit refreshes after replacing either workbook.
"""

from app import app, import_inventory


if __name__ == "__main__":
    with app.app_context():
        import_inventory()
    print("LabVault inventory import completed.")