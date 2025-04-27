import tkinter as tk
from tkinter import ttk, messagebox, filedialog # Import filedialog for saving
import tkinter.font as tkfont
import matplotlib.pyplot as plt # Import matplotlib for potential future graphing
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg # For embedding plots
import csv # Import csv for saving to file
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Tuple, Any
import io # For specific IO errors
import traceback # Import for debugging tracebacks
import copy # Import copy for storing results if needed (though direct assignment is fine here)

# -----------------------------
# CONSTANTS
# -----------------------------
# Default Input Values
DEFAULT_INSTRUMENT_NAME: str = "ETF Investment"
DEFAULT_INITIAL_SUM: str = "10000"
DEFAULT_PERIOD: str = "20"
DEFAULT_GROWTH_RATE: str = "7"
DEFAULT_DIVIDEND_YIELD: str = "2"

# Reinvestment Options (Constants are no longer directly used for *selecting* a scenario)
# REINVEST_ALL = 1
# REINVEST_50_PERCENT = 2
# REINVEST_NONE = 3

# -----------------------------
# DATA STRUCTURE FOR RESULTS
# -----------------------------
@dataclass
class InvestmentYearResult:
    Year: int
    StartingValue: float # Starting value for the year (using 'All' as reference)
    CapitalGrowth: float # Based on StartingValue of the 'All' strategy
    GrossDividend: float # Based on StartingValue of the 'All' strategy

    # Results for All Reinvested (100%)
    Gain_100: float # Renamed from Gain_All for consistency
    EndValue_100: float # Renamed from EndValue_All
    CumulativeGrossDividend_100: float # Renamed from CumulativeGrossDividend_All

    # Results for 75% Reinvested
    # Reinvested_75: float # Can be derived if needed
    Withdrawn_75: float # Yearly withdrawn amount
    CumulativeWithdrawn_75: float # Cumulative withdrawn amount
    Gain_75: float
    EndValue_75: float

    # Results for 50% Reinvested
    # Reinvested_50: float # Can be derived if needed
    Withdrawn_50: float # Yearly withdrawn amount
    CumulativeWithdrawn_50: float # Cumulative withdrawn amount
    Gain_50: float
    EndValue_50: float

    # Results for 25% Reinvested
    # Reinvested_25: float # Can be derived if needed
    Withdrawn_25: float # Yearly withdrawn amount
    CumulativeWithdrawn_25: float # Cumulative withdrawn amount
    Gain_25: float
    EndValue_25: float

    # Results for 0% Reinvested
    # Reinvested_0: float # Can be derived if needed
    Withdrawn_0: float # This will be the full dividend (Yearly amount)
    CumulativeWithdrawn_0: float # Cumulative withdrawn amount
    Gain_0: float      # This will be just Capital Growth
    EndValue_0: float


# Type alias for clarity
SimulationResults = List[InvestmentYearResult]

# -----------------------------
# SIMULATION LOGIC (Integrated into calculate method)
# -----------------------------
# The main simulation loop resides in the calculate method to compute all 5 strategies side-by-side.


# -----------------------------
# GUI USING TKINTER
# -----------------------------
class InvestmentGrowthGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Investment Growth Calculator")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        # Stores the results after a successful calculation for saving
        self.simulation_results: SimulationResults = [] # Initialize to empty list

        # --- Main Frame ---
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.rowconfigure(1, weight=1) # Output frame should expand vertically
        self.main_frame.columnconfigure(0, weight=1) # Input/Output frames expand horizontally

        # --- Input Frame ---
        self.input_frame = ttk.LabelFrame(self.main_frame, text="Inputs", padding="10")
        self.input_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.input_frame.columnconfigure(1, weight=1) # Make entry column expandable

        # --- Output Frame ---
        self.output_frame = ttk.LabelFrame(self.main_frame, text="Results", padding="10")
        self.output_frame.grid(row=1, column=0, padx=5, pady=(5, 5), sticky="nsew")
        self.output_frame.rowconfigure(0, weight=1) # Text area row
        self.output_frame.columnconfigure(0, weight=1) # Text area column

        # --- Tkinter Variables for Input Fields ---
        self.sv_instrument_name = tk.StringVar(value=DEFAULT_INSTRUMENT_NAME)
        self.sv_initial_sum = tk.StringVar(value=DEFAULT_INITIAL_SUM)
        self.sv_period = tk.StringVar(value=DEFAULT_PERIOD)
        self.sv_growth_rate = tk.StringVar(value=DEFAULT_GROWTH_RATE)
        self.sv_dividend_yield = tk.StringVar(value=DEFAULT_DIVIDEND_YIELD)

        # Reinvestment Option Variable removed as all strategies are always shown
        # self.reinvestment_var = tk.IntVar(value=REINVEST_ALL) # Default to All Reinvested

        self.create_input_widgets()
        self.create_output_widgets()

    def _add_input_row(self, label_text: str, string_var: tk.StringVar, row: int, is_numeric: bool = True):
        """Helper to add a label and entry row linked to a StringVar."""
        ttk.Label(self.input_frame, text=label_text).grid(row=row, column=0, sticky="w", padx=5, pady=2)
        entry = ttk.Entry(self.input_frame, textvariable=string_var) # Link textvariable
        entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        # Optional: Add validation command if numeric (can be added here)

    def create_input_widgets(self):
        """Creates and lays out the input widgets."""
        row = 0
        self._add_input_row("Instrument Name:", self.sv_instrument_name, row, is_numeric=False); row += 1
        self._add_input_row("Initial Investment Sum (€):", self.sv_initial_sum, row); row += 1
        self._add_input_row("Investment Period (Years):", self.sv_period, row); row += 1
        self._add_input_row("Expected Yearly Growth (%):", self.sv_growth_rate, row); row += 1
        self._add_input_row("Expected Dividend Yield (%):", self.sv_dividend_yield, row); row += 1

        # Reinvestment Option Radio Buttons removed as all strategies are always shown
        # ttk.Label(self.input_frame, text="Reinvestment Strategy:").grid(row=row, column=0, sticky="w", padx=5, pady=5)
        # rb_frame = ttk.Frame(self.input_frame)
        # rb_frame.grid(row=row, column=1, sticky="w")
        # ttk.Radiobutton(rb_frame, text="Reinvest All Dividends", variable=self.reinvestment_var, value=REINVEST_ALL).pack(anchor="w")
        # ttk.Radiobutton(rb_frame, text="Reinvest 50% of Dividends", variable=self.reinvestment_var, value=REINVEST_50_PERCENT).pack(anchor="w")
        # ttk.Radiobutton(rb_frame, text="Reinvest 0% of Dividends", variable=self.reinvestment_var, value=REINVEST_NONE).pack(anchor="w")
        # row += 1 # Increment row even if radio buttons are removed, to keep layout consistent if other items follow


        # Action Button Frame
        action_frame = ttk.Frame(self.input_frame)
        # Adjust row index if radio buttons were removed and nothing else was added above this
        # Assuming the next available row after input fields is where buttons start.
        # The last _add_input_row increments row, so buttons are on 'row'.
        action_frame.grid(row=row, column=0, columnspan=2, pady=10) # Use the current row index


        self.btn_calculate = ttk.Button(action_frame, text="Calculate Growth", command=self.calculate)
        self.btn_calculate.pack(side=tk.LEFT, padx=5)
        self.btn_new_calc = ttk.Button(action_frame, text="New Calculation", command=self.new_calculation)
        self.btn_new_calc.pack(side=tk.LEFT, padx=5)
        # Add the save button
        self.btn_save_results = ttk.Button(action_frame, text="Save Results", command=self.save_results_ui, state=tk.DISABLED) # Disabled initially
        self.btn_save_results.pack(side=tk.LEFT, padx=5)


    def create_output_widgets(self):
        """Creates the text area for output."""
        # Text Results Area with Scrollbar
        text_frame = ttk.Frame(self.output_frame)
        text_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self.text_results = tk.Text(text_frame, height=15, width=100, wrap="none", borderwidth=0) # Use wrap="none" for tables
        text_scrollbar_y = ttk.Scrollbar(text_frame, orient="vertical", command=self.text_results.yview)
        text_scrollbar_x = ttk.Scrollbar(text_frame, orient="horizontal", command=self.text_results.xview)
        self.text_results.configure(yscrollcommand=text_scrollbar_y.set, xscrollcommand=text_scrollbar_x.set, state='disabled')

        self.text_results.grid(row=0, column=0, sticky="nsew")
        text_scrollbar_y.grid(row=0, column=1, sticky="ns")
        text_scrollbar_x.grid(row=1, column=0, sticky="ew")

        # Placeholder for graph frame if adding plots later
        # self.graph_frame = ttk.Frame(self.output_frame)
        # self.graph_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)


    def _get_float_input(self, string_var: tk.StringVar, label_text: str) -> float:
        """Safely gets and converts value from a StringVar to float, providing label text for errors."""
        try:
            value_str = string_var.get().replace(',', '.') # Handle comma as decimal separator
            return float(value_str)
        except ValueError:
            raise ValueError(f"Vigane sisend väljal '{label_text}': Palun sisesta number.")
        except Exception as e:
             raise ValueError(f"Viga sisendi lugemisel väljal '{label_text}': {e}")

    def _get_int_input(self, string_var: tk.StringVar, label_text: str) -> int:
        """Safely gets and converts value from a StringVar to integer."""
        try:
            value_str = string_var.get().replace(',', '.') # Handle comma, then int will truncate
            return int(float(value_str)) # Convert via float to handle decimals before int conversion
        except ValueError:
            raise ValueError(f"Vigane sisend väljal '{label_text}': Palun sisesta täisarv.")
        except Exception as e:
             raise ValueError(f"Viga sisendi lugemisel väljal '{label_text}': {e}")


    def calculate(self):
        """Reads inputs, performs simulation, and displays results."""
        self.new_calculation() # Clear previous results

        try:
            # Read inputs
            instrument_name = self.sv_instrument_name.get()
            initial_sum = self._get_float_input(self.sv_initial_sum, "Initial Investment Sum")
            period_years = self._get_int_input(self.sv_period, "Investment Period")
            growth_rate = self._get_float_input(self.sv_growth_rate, "Expected Yearly Growth")
            dividend_yield = self._get_float_input(self.sv_dividend_yield, "Expected Dividend Yield")
            # Reinvestment option is no longer read from a variable as all are calculated

            # Basic Validation
            if initial_sum < 0: raise ValueError("Initial Investment Sum cannot be negative.")
            if period_years <= 0: raise ValueError("Investment Period must be a positive integer.")
            if growth_rate < -100: raise ValueError("Growth Rate cannot be less than -100%.")
            if dividend_yield < 0: raise ValueError("Dividend Yield cannot be negative.")

        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
            return
        except Exception as e:
            messagebox.showerror("Error reading inputs", f"An unexpected error occurred: {e}")
            traceback.print_exc()
            return

        # --- Simulation Logic (Calculating all 5 strategies in parallel) ---
        results: SimulationResults = []
        current_value_100 = initial_sum # Renamed from _all
        current_value_75 = initial_sum # New
        current_value_50 = initial_sum
        current_value_25 = initial_sum # New
        current_value_0 = initial_sum

        # Initialize cumulative amounts
        cumulative_gross_dividend_100 = 0.0 # Renamed
        cumulative_withdrawn_75 = 0.0 # New
        cumulative_withdrawn_50 = 0.0
        cumulative_withdrawn_25 = 0.0 # New
        cumulative_withdrawn_0 = 0.0

        try:
            for year in range(1, period_years + 1):
                # Starting values for THIS year for each strategy
                start_value_100 = current_value_100
                start_value_75 = current_value_75
                start_value_50 = current_value_50
                start_value_25 = current_value_25
                start_value_0 = current_value_0

                # Calculate Capital Growth and Gross Dividend for THIS year based on each strategy's starting value
                # This approach correctly reflects compounding differences on both capital and dividends.
                # We calculate these based on the starting value of the 'All Reinvested' (100%) strategy
                # for consistency in the shared table columns (Cap Growth, Gross Div).
                # The actual gain/end value for each strategy uses its *own* start value for accuracy.
                cap_growth_base = start_value_100 * (growth_rate / 100.0)
                gross_div_base = start_value_100 * (dividend_yield / 100.0)

                # --- Calculate for 100% Reinvested ---
                # Use start_value_100 for accurate compounding
                cap_growth_100_actual = start_value_100 * (growth_rate / 100.0)
                gross_div_100_actual = start_value_100 * (dividend_yield / 100.0)
                reinvested_100 = gross_div_100_actual # All is reinvested
                gain_100 = cap_growth_100_actual + reinvested_100
                end_value_100 = start_value_100 + gain_100


                # --- Calculate for 75% Reinvested ---
                # Use start_value_75 for accurate compounding
                cap_growth_75_actual = start_value_75 * (growth_rate / 100.0)
                gross_div_75_actual = start_value_75 * (dividend_yield / 100.0)
                reinvested_75 = gross_div_75_actual * 0.75
                withdrawn_75_yearly = gross_div_75_actual * 0.25
                gain_75 = cap_growth_75_actual + reinvested_75
                end_value_75 = start_value_75 + gain_75


                # --- Calculate for 50% Reinvested ---
                # Use start_value_50 for accurate compounding
                cap_growth_50_actual = start_value_50 * (growth_rate / 100.0)
                gross_div_50_actual = start_value_50 * (dividend_yield / 100.0)
                reinvested_50 = gross_div_50_actual * 0.50
                withdrawn_50_yearly = gross_div_50_actual * 0.50
                gain_50 = cap_growth_50_actual + reinvested_50
                end_value_50 = start_value_50 + gain_50


                # --- Calculate for 25% Reinvested ---
                # Use start_value_25 for accurate compounding
                cap_growth_25_actual = start_value_25 * (growth_rate / 100.0)
                gross_div_25_actual = start_value_25 * (dividend_yield / 100.0)
                reinvested_25 = gross_div_25_actual * 0.25
                withdrawn_25_yearly = gross_div_25_actual * 0.75
                gain_25 = cap_growth_25_actual + reinvested_25
                end_value_25 = start_value_25 + gain_25


                # --- Calculate for 0% Reinvested ---
                # Use start_value_0 for accurate compounding
                cap_growth_0_actual = start_value_0 * (growth_rate / 100.0)
                gross_div_0_actual = start_value_0 * (dividend_yield / 100.0)
                reinvested_0 = gross_div_0_actual * 0.0 # Nothing is reinvested
                withdrawn_0_yearly = gross_div_0_actual * 1.0 # Full dividend is withdrawn
                gain_0 = cap_growth_0_actual + reinvested_0 # Gain is just capital growth
                end_value_0 = start_value_0 + gain_0


                # --- Update Cumulative Amounts ---
                # Cumulative gross div for 100% strategy
                cumulative_gross_dividend_100 += gross_div_100_actual

                # Cumulative withdrawn for other strategies
                cumulative_withdrawn_75 += withdrawn_75_yearly
                cumulative_withdrawn_50 += withdrawn_50_yearly
                cumulative_withdrawn_25 += withdrawn_25_yearly
                cumulative_withdrawn_0 += withdrawn_0_yearly
                # --- End Update ---


                # Store results for this year
                results.append(InvestmentYearResult(
                    Year=year,
                    StartingValue=start_value_100, # Use 100% starting value for this column as a reference
                    CapitalGrowth=cap_growth_100_actual, # Use 100% capital growth as a reference
                    GrossDividend=gross_div_100_actual, # Use 100% gross dividend as a reference

                    Gain_100=gain_100,
                    EndValue_100=end_value_100,
                    CumulativeGrossDividend_100=cumulative_gross_dividend_100,

                    Withdrawn_75=withdrawn_75_yearly,
                    CumulativeWithdrawn_75=cumulative_withdrawn_75,
                    Gain_75=gain_75,
                    EndValue_75=end_value_75,

                    Withdrawn_50=withdrawn_50_yearly,
                    CumulativeWithdrawn_50=cumulative_withdrawn_50,
                    Gain_50=gain_50,
                    EndValue_50=end_value_50,

                    Withdrawn_25=withdrawn_25_yearly,
                    CumulativeWithdrawn_25=cumulative_withdrawn_25,
                    Gain_25=gain_25,
                    EndValue_25=end_value_25,

                    Withdrawn_0=withdrawn_0_yearly,
                    CumulativeWithdrawn_0=cumulative_withdrawn_0,
                    Gain_0=gain_0,
                    EndValue_0=end_value_0
                ))

                # Update current values for the next year's start
                current_value_100 = end_value_100
                current_value_75 = end_value_75
                current_value_50 = end_value_50
                current_value_25 = end_value_25
                current_value_0 = end_value_0

        except Exception as e:
            messagebox.showerror("Simulation Error", f"An error occurred during simulation: {e}")
            traceback.print_exc()
            return

        # --- Store results for saving ---
        self.simulation_results = results # Store the list of InvestmentYearResult objects
        # --- End Store ---

        # --- Display Results ---
        self.display_results(instrument_name, initial_sum, period_years, growth_rate, dividend_yield, results)

        # --- Enable Save Button ---
        # Assumes self.btn_save_results is an attribute created in create_input_widgets
        if hasattr(self, 'btn_save_results'):
             self.btn_save_results.config(state=tk.NORMAL)
        # --- End Enable ---


    def display_results(self, instrument_name: str, initial_sum: float, period_years: int, growth_rate: float, dividend_yield: float, results: SimulationResults):
        """Formats and displays simulation results in the text area."""
        output_text = f"--- Simulation Results: {instrument_name} ---\n"
        output_text += f"Initial Investment: {initial_sum:,.2f} €\n".replace(",", " ")
        output_text += f"Period: {period_years} years\n"
        output_text += f"Expected Yearly Growth: {growth_rate:,.2f} %\n".replace(",", " ")
        output_text += f"Expected Dividend Yield: {dividend_yield:,.2f} %\n".replace(",", " ")
        output_text += "-" * 50 + "\n\n"

        # Define table headers (Adding 75% and 25% columns)
        headers = [
            "Year", "Start Value", "Cap Growth", "Gross Div", # Shared (4 columns)
            "Gain (100%)", "End Value (100%)", "Cum. Gross Div (100%)", # 100% (3 columns)
            "Gain (75%)", "End Value (75%)", "Withdrawn (75%)", "Cum. Withdrawn (75%)", # 75% (4 columns)
            "Gain (50%)", "End Value (50%)", "Withdrawn (50%)", "Cum. Withdrawn (50%)", # 50% (4 columns)
            "Gain (25%)", "End Value (25%)", "Withdrawn (25%)", "Cum. Withdrawn (25%)", # 25% (4 columns)
            "Gain (0%)", "End Value (0%)", "Withdrawn (0%)", "Cum. Withdrawn (0%)" # 0% (4 columns)
        ] # Total columns: 4 + 3 + 4*4 = 7 + 16 = 23 columns


        # Define format string for alignment (adjust spacing as needed for 23 columns)
        # Example: {:<5} {:>15} {:>15} {:>15} {:>15} {:>18} {:>18} {:>15} {:>18} {:>15} {:>18} {:>15} {:>18} {:>15} {:>18} {:>15} {:>18} {:>15} {:>18} {:>15} {:>18} {:>15} {:>18}\n
        # Let's refine spacing slightly for 23 columns
        fmt = "{:<5} {:>12} {:>12} {:>12} {:>12} {:>15} {:>15} {:>12} {:>15} {:>12} {:>15} {:>12} {:>15} {:>12} {:>15} {:>12} {:>15} {:>12} {:>15} {:>12} {:>15} {:>12} {:>15}\n"

        # Add headers and separator
        output_text += fmt.format(*headers)
        separator = "-" * (len(fmt.format(*headers).rstrip())) + "\n"
        output_text += separator

        # Add yearly results (Update formatting for new columns)
        for r in results:
            output_text += fmt.format(
                r.Year,
                f"{r.StartingValue:,.0f}".replace(",", " "),
                f"{r.CapitalGrowth:,.0f}".replace(",", " "),
                f"{r.GrossDividend:,.0f}".replace(",", " "),

                f"{r.Gain_100:,.0f}".replace(",", " "),
                f"{r.EndValue_100:,.0f}".replace(",", " "),
                f"{r.CumulativeGrossDividend_100:,.0f}".replace(",", " "),

                f"{r.Gain_75:,.0f}".replace(",", " "),
                f"{r.EndValue_75:,.0f}".replace(",", " "),
                f"{r.Withdrawn_75:,.0f}".replace(",", " "),
                f"{r.CumulativeWithdrawn_75:,.0f}".replace(",", " "),

                f"{r.Gain_50:,.0f}".replace(",", " "),
                f"{r.EndValue_50:,.0f}".replace(",", " "),
                f"{r.Withdrawn_50:,.0f}".replace(",", " "),
                f"{r.CumulativeWithdrawn_50:,.0f}".replace(",", " "),

                f"{r.Gain_25:,.0f}".replace(",", " "),
                f"{r.EndValue_25:,.0f}".replace(",", " "),
                f"{r.Withdrawn_25:,.0f}".replace(",", " "),
                f"{r.CumulativeWithdrawn_25:,.0f}".replace(",", " "),

                f"{r.Gain_0:,.0f}".replace(",", " "),
                f"{r.EndValue_0:,.0f}".replace(",", " "),
                f"{r.Withdrawn_0:,.0f}".replace(",", " "),
                f"{r.CumulativeWithdrawn_0:,.0f}".replace(",", " ")
            )

        # Add separator and summary row (Update formatting for all 23 columns)
        output_text += separator
        if results:
            last_year = results[-1]
            summary_label = f"End ({period_years} Yrs)"
            output_text += fmt.format(
                summary_label,
                "", # Year
                "", # Start Value
                "", # Capital Growth
                "", # Gross Dividend
                f"{last_year.Gain_100:,.0f}".replace(",", " "), # Last year's gain (100%)
                f"{last_year.EndValue_100:,.0f}".replace(",", " "), # Final End Value 100%
                f"{last_year.CumulativeGrossDividend_100:,.0f}".replace(",", " "), # Final Cumulative Gross Div 100%

                f"{last_year.Gain_75:,.0f}".replace(",", " "), # Last year's gain (75%)
                f"{last_year.EndValue_75:,.0f}".replace(",", " "), # Final End Value 75%
                f"{last_year.Withdrawn_75:,.0f}".replace(",", " "), # Last year's withdrawn (75%)
                f"{last_year.CumulativeWithdrawn_75:,.0f}".replace(",", " "), # Final Cumulative 75%

                 f"{last_year.Gain_50:,.0f}".replace(",", " "), # Last year's gain (50%)
                f"{last_year.EndValue_50:,.0f}".replace(",", " "), # Final End Value 50%
                f"{last_year.Withdrawn_50:,.0f}".replace(",", " "), # Last year's withdrawn (50%)
                f"{last_year.CumulativeWithdrawn_50:,.0f}".replace(",", " "), # Final Cumulative 50%

                 f"{last_year.Gain_25:,.0f}".replace(",", " "), # Last year's gain (25%)
                f"{last_year.EndValue_25:,.0f}".replace(",", " "), # Final End Value 25%
                f"{last_year.Withdrawn_25:,.0f}".replace(",", " "), # Last year's withdrawn (25%)
                f"{last_year.CumulativeWithdrawn_25:,.0f}".replace(",", " "), # Final Cumulative 25%

                 f"{last_year.Gain_0:,.0f}".replace(",", " "), # Last year's gain (0%)
                f"{last_year.EndValue_0:,.0f}".replace(",", " "), # Final End Value 0%
                f"{last_year.Withdrawn_0:,.0f}".replace(",", " "), # Last year's withdrawn (0%)
                f"{last_year.CumulativeWithdrawn_0:,.0f}".replace(",", " ") # Final Cumulative 0%
            )

            # Add extra summary lines specifically for total cumulative amounts
            output_text += "\n"
            output_text += "Total Cumulative Amounts Over Period:\n"
            output_text += f"  100% Reinvested (Gross Dividends Generated): {last_year.CumulativeGrossDividend_100:,.2f} €\n".replace(",", " ")
            output_text += f"  75% Reinvestment (Dividends Withdrawn): {last_year.CumulativeWithdrawn_75:,.2f} €\n".replace(",", " ")
            output_text += f"  50% Reinvestment (Dividends Withdrawn): {last_year.CumulativeWithdrawn_50:,.2f} €\n".replace(",", " ")
            output_text += f"  25% Reinvestment (Dividends Withdrawn): {last_year.CumulativeWithdrawn_25:,.2f} €\n".replace(",", " ")
            output_text += f"  0% Reinvestment (Dividends Withdrawn): {last_year.CumulativeWithdrawn_0:,.2f} €\n".replace(",", " ")


        # Clear previous results before inserting new ones
        self.text_results.configure(state='normal') # Ensure text widget is editable
        self.text_results.delete("1.0", tk.END)
        self.text_results.insert(tk.END, output_text)
        self.text_results.configure(state='disabled') # Make read-only after update


    def new_calculation(self):
        """Clears the output area and stored results."""
        self.text_results.configure(state='normal')
        self.text_results.delete("1.0", tk.END)
        self.text_results.configure(state='disabled')
        self.simulation_results = [] # Clear stored results
        # Disable save button if it exists
        if hasattr(self, 'btn_save_results'):
             self.btn_save_results.config(state=tk.DISABLED)

    def save_results_ui(self):
        """
        Prompts user for filename and saves full simulation results for all
        strategies to a single CSV file. Includes a 'Total' row.
        """
        if not self.simulation_results:
             messagebox.showwarning("Salvestamine", "Arvutustulemused puuduvad. Palun käivita 'Calculate Growth' enne salvestamist.")
             return

        initial_filename = "investment_growth_simulation.csv"

        file_path = filedialog.asksaveasfilename(
            initialfile=initial_filename,
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Salvesta simulatsiooni tulemused"
        )

        if not file_path:
            # User cancelled
            return

        try:
            # Define fieldnames based on the InvestmentYearResult dataclass fields
            # Use the asdict keys, which match the dataclass field names
            fieldnames = list(asdict(self.simulation_results[0]).keys()) if self.simulation_results else []

            if not fieldnames:
                 messagebox.showwarning("Salvestamine", "No data available to save.")
                 return

            # Use DictWriter, handle missing keys gracefully with default ""
            # Use semicolon as delimiter for compatibility with European locales
            with open(file_path, mode="w", newline="", encoding="utf-8-sig") as csvfile: # utf-8-sig for Excel BOM
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames, restval="", extrasaction='ignore', delimiter=';')
                writer.writeheader()

                # Write each yearly result row
                for row_data in self.simulation_results:
                    row_dict = asdict(row_data)
                     # Round numeric values intended for CSV output (optional, but matches text display)
                    for key, value in row_dict.items():
                         if isinstance(value, (int, float)):
                             row_dict[key] = round(value, 2) # Round to 2 decimal places for CSV

                    writer.writerow(row_dict)

                # Add a summary row (similar data as the last row, but maybe labeled)
                # This provides a clear "Summary" row header in CSV.
                if self.simulation_results:
                    last_year_results = asdict(self.simulation_results[-1])
                    summary_row = {key: "" for key in fieldnames} # Start with empty row
                    summary_row["Year"] = f"Summary ({last_year_results.get('Year', '')} Yrs)" # Label the row using last year's year

                    # Manually copy key final values from the last year's data for the summary row
                    # This ensures only relevant summary data appears under the 'Summary' row header
                    summary_fields_to_copy = [
                        "EndValue_100", "CumulativeGrossDividend_100",
                        "EndValue_75", "CumulativeWithdrawn_75",
                        "EndValue_50", "CumulativeWithdrawn_50",
                        "EndValue_25", "CumulativeWithdrawn_25",
                        "EndValue_0", "CumulativeWithdrawn_0",
                        # Include last year's yearly withdrawn for context if desired
                        "Withdrawn_75", "Withdrawn_50", "Withdrawn_25", "Withdrawn_0",
                         # Include last year's gain for context if desired
                         "Gain_100", "Gain_75", "Gain_50", "Gain_25", "Gain_0",
                          # Include last year's shared values for context if desired
                         "StartingValue", "CapitalGrowth", "GrossDividend"

                    ]

                    for field_name in fieldnames:
                         if field_name in summary_fields_to_copy:
                             value = last_year_results.get(field_name)
                             if isinstance(value, (int, float)):
                                 summary_row[field_name] = round(value, 2)
                             else:
                                  summary_row[field_name] = value # Copy non-numeric as is

                    writer.writerow(summary_row)


            # Add confirmation to text widget and show message box
            self.text_results.configure(state='normal')
            self.text_results.insert(tk.END, f"\nTulemused salvestatud faili:\n{file_path}\n")
            self.text_results.configure(state='disabled')
            messagebox.showinfo("Salvestamine õnnestus", f"Tulemused salvestati faili:\n{file_path}")

        except (IOError, PermissionError, csv.Error) as e:
            error_msg = f"Salvestamise viga faili {file_path}:\n{e}\n\nVeendu, et fail pole avatud teises programmis ja sul on kausta kirjutusõigus."
            messagebox.showerror("Salvestamise Viga", error_msg)
        except Exception as e: # Catch any other unexpected errors
            error_msg = f"Ootamatu viga salvestamisel:\n{e}"
            traceback.print_exc() # Log for debugging
            messagebox.showerror("Salvestamise Viga", error_msg)


# -----------------------------
# MAIN APPLICATION ENTRY POINT
# -----------------------------
if __name__ == "__main__":
    root = tk.Tk()

    # --- Style and Font Configuration ---
    style = ttk.Style()
    try:
        if 'clam' in style.theme_names():
            style.theme_use('clam')
        elif 'vista' in style.theme_names(): # Windows
             style.theme_use('vista')
        elif 'aqua' in style.theme_names(): # macOS
             style.theme_use('aqua')
        # else fallback to default
    except tk.TclError:
        print("Could not set preferred theme, using default.")

    default_font = tkfont.nametofont("TkDefaultFont")
    font_size = default_font.cget("size")
    new_size = max(10, int(font_size * 1.1)) # Ensure minimum size 10
    default_font.configure(size=new_size)
    root.option_add("*Font", default_font)
    style.configure("TLabel", font=default_font)
    style.configure("TButton", font=default_font)
    style.configure("TCheckbutton", font=default_font)
    style.configure("TRadiobutton", font=default_font)
    style.configure("TEntry", font=default_font)
    style.configure("TLabelFrame.Label", font=default_font) # For LabelFrame titles


    root.minsize(800, 600) # Increase min size for wider table

    app = InvestmentGrowthGUI(root)
    root.mainloop()