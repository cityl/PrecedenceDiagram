import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import networkx as nx
import pandas as pd
from tkinter import *

class ExcelDiagramApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Diagram Generator")
        self.geometry("800x600")  # Initial window size

        base_font_size = 20  # You can increase this value to scale up the text size
        self.custom_font = tk.font.Font(family="Helvetica", size=base_font_size)

        # Create a drop target for Excel files
        self.drop_target = tk.Label(self, text="Drag and drop an Excel file here", font=self.custom_font, padx=10, pady=10, relief=tk.SUNKEN)
        self.drop_target.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Bind the drop event to the file_handler method
        self.drop_target.drop_target_register(DND_FILES)
        self.drop_target.dnd_bind('<<Drop>>', self.file_handler)

        # Initialize the Matplotlib figure
        self.figure, self.ax = plt.subplots(figsize=(6, 4))
        self.ax.axis('off')  # Turn off the axis for a cleaner display

        # Create a Matplotlib canvas
        self.canvas = FigureCanvasTkAgg(self.figure, self)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Create a "Show Diagram" button
        self.show_button = Button(self, text="Show Diagram", font=self.custom_font, command=self.show_diagram)
        self.show_button.pack()

        # Initialize graph data
        self.G_int = None
        self.positions_int = None

        # Bind the close event to a custom handler
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def on_close(self):
        # Place any cleanup or termination code here

        # Finally, destroy the window
        self.destroy()

    def file_handler(self, event):
        # Get the file path from the dropped files
        file_path = event.data

        # Read the Excel file into a DataFrame
        try:
            df = pd.read_excel(file_path)

            # Convert the 'Preceded_by' column to a string and then to a list of integers, if not NaN
            df['Preceded_by'] = df['Preceded_by'].fillna('').astype(str)
            df['Preceded_by'] = df['Preceded_by'].apply(lambda x: [self.safe_int_conversion(i.strip()) for i in x.split(',') if i.strip().isdigit()])

            # Create the tasks_table dictionary from the DataFrame
            tasks_table = df.set_index('Element')['Preceded_by'].to_dict()

            # Convert node labels in the tasks_table to integers where possible
            tasks_table_int = {self.safe_int_conversion(task): [self.safe_int_conversion(dep) for dep in deps] for task, deps in tasks_table.items()}

            # Recalculate levels and positions with the updated tasks_table
            levels_int = self.assign_levels(tasks_table_int)
            positions_int = self.create_positions(levels_int)

            # Recreate the directed graph with the updated tasks_table
            self.G_int = nx.DiGraph()
            for task, dependencies in tasks_table_int.items():
                for dependency in dependencies:
                    self.G_int.add_edge(dependency, task)

            # Clear the existing figure
            self.ax.clear()

            # Draw the graph with integer labels (where possible)
            nx.draw(self.G_int, positions_int, with_labels=True, node_size=1000, node_color='skyblue', font_size=14, font_weight='bold', arrowsize=20, ax=self.ax)
            self.ax.set_title('Precedence Diagram')

            # Store the graph data
            self.positions_int = positions_int

            # Refresh the canvas
            self.canvas.draw()

        except Exception as e:
            # Handle any exceptions that may occur during file reading or diagram generation
            error_message = f"Error: {str(e)}"
            print(error_message)

    def safe_int_conversion(self, value):
        try:
            return int(value)
        except ValueError:
            return value

    def assign_levels(self, tasks_table):
        levels = {task: 0 for task in tasks_table if not tasks_table[task]}
        level = 0
        while len(levels) < len(tasks_table):
            next_level_tasks = [task for task, deps in tasks_table.items() if task not in levels and all(dep in levels for dep in deps)]
            level += 1
            for task in next_level_tasks:
                levels[task] = level
        return levels

    def create_positions(self, levels):
        level_widths = max(levels.values()) + 1
        positions = {}
        for level in range(level_widths):
            tasks_at_level = [task for task, task_level in levels.items() if task_level == level]
            num_tasks = len(tasks_at_level)
            middle_line = (num_tasks) / 2.0
            for i, task in enumerate(sorted(tasks_at_level)):
                positions[task] = (level, middle_line - i)
        return positions

    def show_diagram(self):
        if self.G_int is not None and self.positions_int is not None:
            plt.figure(figsize=(8, 6))
            plt.axis('off')  # Turn off the axis for a cleaner display
            plt.title('Precedence Diagram')
            nx.draw(self.G_int, self.positions_int, with_labels=True, node_size=3000, node_color='skyblue', font_size=28, font_weight='bold', arrowsize=40)
            plt.show()

if __name__ == "__main__":
    app = ExcelDiagramApp()
    app.mainloop()
