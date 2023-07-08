import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.tree import DecisionTreeClassifier
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
import tkinter.ttk as ttk
from sklearn import tree

class DataAnalysisGUI:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Data Analysis GUI")
        self.datatable = pd.DataFrame()
        self.target_variable = None
        self.column_filters = {}
        self.selected_columns = []
        self.create_menu()
        self.create_toolbar()
        self.create_column_selection()
        self.create_datatable()
        self.create_statistical_output()
        self.create_filter_entries()
        self.create_buttons()
        self.create_checkboxes()

    def run(self):
        self.root.geometry("800x720")  # Set the initial size of the window
        self.root.mainloop()

    def load_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls *.xlsx"), ("CSV Files", "*.csv")])
        if file_path:
            try:
                if file_path.endswith((".xls", ".xlsx")):
                    self.datatable = pd.read_excel(file_path)
                elif file_path.endswith(".csv"):
                    self.datatable = pd.read_csv(file_path)
                self.update_datatable()
                self.update_column_listbox()
            except Exception as e:
                messagebox.showerror("Load Data Error", f"Error occurred while loading data:\n{str(e)}")

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open", command=self.load_data)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)

    def create_toolbar(self):
        toolbar_frame = tk.Frame(self.root)
        toolbar_frame.pack(fill=tk.X)

        load_button = tk.Button(toolbar_frame, text="Load Data", command=self.load_data)
        load_button.pack(side=tk.LEFT, padx=5, pady=5)

    def create_datatable(self):
        self.datatable_frame = tk.Frame(self.root)
        self.datatable_frame.pack(fill=tk.BOTH, expand=True)

        self.datatable_treeview = ttk.Treeview(self.datatable_frame)
        self.datatable_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.datatable_scrollbar = ttk.Scrollbar(self.datatable_frame, orient=tk.VERTICAL,
                                                 command=self.datatable_treeview.yview)
        self.datatable_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.datatable_treeview.configure(yscrollcommand=self.datatable_scrollbar.set)

    def create_statistical_output(self):
        stat_frame = tk.LabelFrame(self.root, text="Statistical Output")
        stat_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.stat_text = tk.Text(stat_frame, width=80, height=10)
        self.stat_text.pack(fill=tk.BOTH, expand=True)

    def create_buttons(self):
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X)

        stat_button = tk.Button(button_frame, text="Calculate Statistics", command=self.calculate_statistics)
        vis_button = tk.Button(button_frame, text="Visualize Data", command=self.visualize_data)
        ml_button = tk.Button(button_frame, text="Run Decision Tree", command=self.run_decision_tree)
        linear_regression_button = tk.Button(button_frame, text="Linear Regression", command=self.run_linear_regression)
        random_forest_button = tk.Button(button_frame, text="Random Forest", command=self.run_random_forest)
        save_button = tk.Button(button_frame, text="Save Results", command=self.save_results)

        stat_button.pack(side=tk.LEFT, padx=5, pady=5)
        vis_button.pack(side=tk.LEFT, padx=5, pady=5)
        ml_button.pack(side=tk.LEFT, padx=5, pady=5)
        linear_regression_button.pack(side=tk.LEFT, padx=5, pady=5)
        random_forest_button.pack(side=tk.LEFT, padx=5, pady=5)
        save_button.pack(side=tk.LEFT, padx=5, pady=5)

    def create_checkboxes(self):
        checkbox_frame = tk.Frame(self.root)
        checkbox_frame.pack(fill=tk.X)

        self.skew_var = tk.IntVar()
        self.kurt_var = tk.IntVar()
        self.mean_var = tk.IntVar()
        self.median_var = tk.IntVar()

        skew_checkbox = tk.Checkbutton(checkbox_frame, text="Skew", variable=self.skew_var)
        kurt_checkbox = tk.Checkbutton(checkbox_frame, text="Kurtosis", variable=self.kurt_var)
        mean_checkbox = tk.Checkbutton(checkbox_frame, text="Mean", variable=self.mean_var)
        median_checkbox = tk.Checkbutton(checkbox_frame, text="Median", variable=self.median_var)

        skew_checkbox.pack(side=tk.LEFT, padx=5, pady=5)
        kurt_checkbox.pack(side=tk.LEFT, padx=5, pady=5)
        mean_checkbox.pack(side=tk.LEFT, padx=5, pady=5)
        median_checkbox.pack(side=tk.LEFT, padx=5, pady=5)

    def create_column_selection(self):
        selection_frame = tk.Frame(self.root)
        selection_frame.pack(fill=tk.X)

        label = tk.Label(selection_frame, text="Select column(s) for visualization:")
        label.pack(side=tk.LEFT, padx=5, pady=5)

        self.column_listbox = tk.Listbox(selection_frame, selectmode=tk.MULTIPLE)
        self.column_listbox.pack(side=tk.LEFT, padx=5, pady=5)
        self.column_listbox.bind('<<ListboxSelect>>', self.on_column_select)

    def create_filter_entries(self):
        filter_frame = tk.Frame(self.root)
        filter_frame.pack(fill=tk.X)

        label = tk.Label(filter_frame, text="Filter Rows:")
        label.pack(side=tk.LEFT, padx=5, pady=5)

        self.filter_entries = {}
        for column in self.datatable.columns:
            label = tk.Label(filter_frame, text=column)
            label.pack(side=tk.LEFT, padx=5, pady=5)

            unique_values = self.datatable[column].unique().tolist()
            combobox = ttk.Combobox(filter_frame, values=["All"] + unique_values, state="readonly")
            combobox.pack(side=tk.LEFT, padx=5, pady=5)

            self.filter_entries[column] = combobox

        apply_button = tk.Button(filter_frame, text="Apply Filters", command=self.apply_filters)
        apply_button.pack(side=tk.LEFT, padx=5, pady=5)


    def apply_row_filters(self):
        filtered_data = self.datatable.copy()

        for column, entry in self.filter_entries.items():
            filter_value = entry.get()
            if filter_value != "All":
                filtered_data = filtered_data[filtered_data[column].astype(str) == filter_value]

        self.update_datatable(filtered_data)

    def update_datatable(self, filtered_rows=None):
        self.datatable_treeview.delete(*self.datatable_treeview.get_children())

        if filtered_rows is None:
            data = self.datatable
        else:
            data = filtered_rows

        columns = list(data.columns)
        self.datatable_treeview["columns"] = columns

        self.datatable_treeview.column("#0", width=0, stretch=tk.NO)

        for column in columns:
            self.datatable_treeview.heading(column, text=column)
            self.datatable_treeview.column(column, anchor=tk.CENTER)

        for _, row in data.iterrows():
            values = list(row.values)
            self.datatable_treeview.insert("", tk.END, values=values)    

    def apply_filters(self):#####
        filtered_data = self.datatable.copy()

        for column, combobox in self.filter_entries.items():
            filter_value = combobox.get()

            if filter_value != "All":
                filtered_data = filtered_data[filtered_data[column].astype(str) == filter_value]

        self.update_datatable(filtered_data)

    def update_column_listbox(self):
        self.column_listbox.delete(0, tk.END)

        for column in self.datatable.columns:
            self.column_listbox.insert(tk.END, column)

    def on_column_select(self, event):
        selected_indices = self.column_listbox.curselection()
        self.selected_columns = [self.column_listbox.get(index) for index in selected_indices]

    def visualize_data(self):
        if len(self.selected_columns) >= 2:
            filtered_data = self.datatable.copy()

            for column, entry in self.filter_entries.items():
                filter_value = entry.get()
                if filter_value != "All":
                    filtered_data = filtered_data[filtered_data[column] == filter_value]

            filtered_data = filtered_data[self.selected_columns]
            if len(filtered_data) > 0:
                filtered_data.plot.scatter(x=self.selected_columns[0], y=self.selected_columns[1])
                plt.xlabel(self.selected_columns[0])
                plt.ylabel(self.selected_columns[1])
                plt.title(f"Scatterplot: {self.selected_columns[0]} vs {self.selected_columns[1]}")
                plt.show()
            else:
                messagebox.showwarning("Data Visualization", "No data available for the selected filters.")
        else:
            messagebox.showwarning("Data Visualization", "Please select at least 2 columns for visualization.")

    def calculate_statistics(self):
        if len(self.selected_columns) >= 2:
            if all(col in self.datatable.columns for col in self.selected_columns):
                stats = {}
                for col in self.selected_columns:
                    col_stats = {}
                    if self.skew_var.get():
                        col_stats['Skew'] = self.datatable[col].skew()
                    if self.kurt_var.get():
                        col_stats['Kurtosis'] = self.datatable[col].kurtosis()
                    if self.mean_var.get():
                        col_stats['Mean'] = self.datatable[col].mean()
                    if self.median_var.get():
                        col_stats['Median'] = self.datatable[col].median()

                    stats[col] = col_stats

                self.stat_text.config(state=tk.NORMAL)
                self.stat_text.delete(1.0, tk.END)
                self.stat_text.insert(tk.END, self.format_statistics(stats))
                self.stat_text.config(state=tk.DISABLED)
            else:
                messagebox.showwarning("Calculate Statistics", "Selected columns do not exist in the dataset.")
        else:
            messagebox.showwarning("Calculate Statistics", "Please select at least 2 columns for analysis.")

    def run_decision_tree(self):
        if len(self.selected_columns) >= 1:
            features = self.datatable[self.selected_columns[:-1]]
            target = self.datatable[self.selected_columns[-1]]
            classifier = DecisionTreeClassifier()
            classifier.fit(features, target)

            plt.figure(figsize=(10, 6))
            tree.plot_tree(classifier, feature_names=features.columns, class_names=target.unique(), filled=True)
            plt.show()

            messagebox.showinfo("Decision Tree", "Classifier trained successfully!")
        else:
            messagebox.showwarning("Run Decision Tree", "Please select at least 1 feature column and 1 target column.")

    def run_linear_regression(self):
        if len(self.selected_columns) >= 1:
            features = self.datatable[self.selected_columns[:-1]]
            target = self.datatable[self.selected_columns[-1]]
            model = LinearRegression()
            model.fit(features, target)

            predictions = model.predict(features)
            plt.scatter(target, predictions)
            plt.xlabel("Actual")
            plt.ylabel("Predicted")
            plt.title("Linear Regression - Actual vs Predicted")
            plt.show()

            messagebox.showinfo("Linear Regression", "Linear regression model trained successfully!")
        else:
            messagebox.showwarning("Run Linear Regression", "Please select at least 1 feature column and 1 target column.")

    def run_random_forest(self):
        if len(self.selected_columns) >= 1:
            features = self.datatable[self.selected_columns[:-1]]
            target = self.datatable[self.selected_columns[-1]]
            model = RandomForestRegressor()
            model.fit(features, target)

            predictions = model.predict(features)
            plt.scatter(target, predictions)
            plt.xlabel("Actual")
            plt.ylabel("Predicted")
            plt.title("Random Forest - Actual vs Predicted")
            plt.show()

            messagebox.showinfo("Random Forest", "Random forest model trained successfully!")
        else:
            messagebox.showwarning("Run Random Forest", "Please select at least 1 feature column and 1 target column.")

    def save_results(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if file_path:
            try:
                self.datatable.to_csv(file_path, index=False)
                messagebox.showinfo("Save Results", "Results saved successfully!")
            except Exception as e:
                messagebox.showerror("Save Results Error", f"Error occurred while saving results:\n{str(e)}")

    def format_statistics(self, stats):
        formatted_stats = ""
        for col, col_stats in stats.items():
            formatted_stats += f"Column: {col}\n"
            for stat, value in col_stats.items():
                formatted_stats += f"{stat}: {value:.2f}\n"
            formatted_stats += "\n"

        return formatted_stats


if __name__ == "__main__":
    app = DataAnalysisGUI()
    app.run()

