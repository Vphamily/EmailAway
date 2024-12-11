import xlwings as xw
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.linear_model import LogisticRegression

class NameCategorizer:
    def __init__(self):
        self.company_clue_words = ["LLC", "INC", "LTD", "CORPORATION", "CORP", "CO.", "LLP", "GMBH", "PTY", "AND",
                                   "COMPANY", "ENTERPRISE", "CONSULTING", "INC.", "TECHNOLOGIES", "COMPA", "CO"]
        self.vectorizer = None
        self.model = None
        self.use_model = False

    def label_name_type(self, name):
        if not name:
            print("Received an empty or None name.")
            return 0  # Treat None or empty as Commercial by default
        name_upper = str(name).upper()
        if any(clue in name_upper for clue in self.company_clue_words):
            return 0  # Commercial
        return 1  # Retail

    def label_commercial_or_retail(self, name):
        if not name:
            return "Commercial"
        label = self.label_name_type(name)
        if self.use_model and label == 0:
            name_vectorized = self.vectorizer.transform([str(name)])
            prediction = self.model.predict(name_vectorized)[0]
            return "Retail" if prediction == 1 else "Commercial"
        else:
            return "Retail" if label == 1 else "Commercial"

    def train_model(self, customer_data):
        cleaned_data = [str(name).strip() for name in customer_data if name]
        if len(cleaned_data) < 10:
            print("Not enough data to enable model training.")
            self.use_model = False
            return
        labels = [self.label_name_type(name) for name in cleaned_data]
        if len(set(labels)) < 2:
            print("Not enough classes to train the model.")
            self.use_model = False
            return

        self.use_model = True
        self.vectorizer = CountVectorizer(stop_words='english')
        X_vectorized = self.vectorizer.fit_transform(cleaned_data)
        X_train, X_test, y_train, y_test = train_test_split(X_vectorized, labels, test_size=0.2, random_state=42)
        self.model = LogisticRegression()
        self.model.fit(X_train, y_train)

    def process_sheet(self, sheet, workbook):
        last_row = sheet.range("A1").end('down').row
        column_data = sheet.range(f"A2:A{last_row}").value
        if column_data is None:
            print("No data found in the specified range.")
            return

        predictions = [self.label_commercial_or_retail(name) for name in column_data if name is not None]
        sheet.range("B1").value = "Category"
        result_range = f"B2:B{len(predictions) + 1}"
        sheet.range(result_range).value = [[pred] for pred in predictions]

def process_file(input_file_path):
    try:
        app = xw.App(visible=False)
        workbook = app.books.open(input_file_path)
        sheet = workbook.sheets[0]
        categorizer = NameCategorizer()
        data = sheet.range("A2:A" + str(sheet.cells.last_cell.row)).value
        categorizer.train_model(data or [])
        categorizer.process_sheet(sheet, workbook)
        workbook.save()
        workbook.close()
        app.quit()
        print(f"Updated file has been saved at: {input_file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Specify the path to the Excel file to be processed
input_file_path = "data.xlsx"  # Ensure this path is correctly set to where the file is

# Call the function to process the file
process_file(input_file_path)
