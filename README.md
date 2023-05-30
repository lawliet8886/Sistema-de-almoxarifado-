**Product Inventory Management System**

This is a simple product inventory management system implemented in Python using PyQt5 for the graphical user interface and SQLite for the database. The system allows users to add, edit, delete, and view products in the inventory, as well as register withdrawals and entries. It also provides the functionality to generate reports of the inventory, withdrawals, and entries data.

**Features**

1. Add a Product: Users can add new products to the inventory by entering the product name, category, quantity, minimum quantity, expiration date, and registration date.

2. Edit a Product: Users can edit the details of an existing product, including its name, category, quantity, minimum quantity, expiration date, and registration date.

3. Delete a Product: Users can delete a product from the inventory. A confirmation prompt ensures that the deletion is intentional.

4. Register Withdrawal: Users can register the withdrawal of a product from the inventory by selecting a product and providing the quantity withdrawn, withdrawal date, and responsible person. The system validates the quantity available in stock before registering the withdrawal.

5. View Withdrawal Details: Users can view the withdrawal details of a specific product, including the quantity withdrawn, withdrawal date, and responsible person.

6. Register Entry: Users can register the entry of additional quantity for a product in the inventory. They need to select a product and provide the quantity entered, entry date, and responsible person.

7. View Entry Details: Users can view the entry details of a specific product, including the quantity entered, entry date, and responsible person.

8. Generate Reports: Users can generate an Excel report that includes information about the products in the inventory, withdrawals, and entries. The report is saved in a specified directory and contains separate worksheets for each category of data.

**Requirements**

To run the application, you need to have the following dependencies installed:

- Python 3.x
- PyQt5
- SQLite3
- openpyxl

You can install these dependencies using pip:

```
pip install pyqt5
pip install openpyxl
```

**Running the Application**

1. Clone the repository or download the source code files.

2. Open a terminal or command prompt and navigate to the directory containing the source code files.

3. Run the following command to start the application:

   ```
   python main.py
   ```

4. The main window of the product inventory management system will appear.

**Usage**

1. Add Products: Click on the "Add" button to add a new product to the inventory. Fill in the required details in the provided input fields and click "Save".

2. Edit Products: Select a product from the inventory table and click the "Edit" button to modify its details. Make the necessary changes in the dialog that appears and click "Save".

3. Delete Products: Select a product from the inventory table and click the "Delete" button to remove it from the inventory. Confirm the deletion when prompted.

4. Register Withdrawal: Select a product from the inventory table and click the "Withdrawal" button to register a product withdrawal. Enter the quantity withdrawn, withdrawal date, and responsible person in the dialog that appears, and click "Save".

5. View Withdrawal Details: Select a product from the inventory table and click the "Withdrawal Details" button to view the withdrawal details for that product. The withdrawal details will be displayed in a dialog window.

6. Register Entry: Select a product from the inventory table and click the "Entry" button to register a product entry. Enter the quantity entered, entry date, and responsible person in the dialog that appears, and click "Save".

7. View Entry Details: Select a product from the inventory table and click the "Entry Details" button to view the entry details for that

 product. The entry details will be displayed in a dialog window.

8. Generate Report: Click the "Generate Report" button to generate an Excel report of the inventory, withdrawals, and entries data. The report will be saved in a directory called "Relatório e Verificação" with the filename "Relatório do Almoxarifado.xlsx".

9. Clear Input Fields: Click the "Clear" button to reset the input fields in the main window.

**Contributing**

Contributions to this product inventory management system are welcome. If you find any bugs or have suggestions for improvements, please open an issue or submit a pull request on the project's GitHub repository.

**License**

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT). Feel free to use and modify the code according to your needs.

**Author**

This product inventory management system was developed by Gabriel da Silva Fernandes

