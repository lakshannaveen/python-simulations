import random
import pandas as pd
import matplotlib.pyplot as plt
from statistics import mean, mode
from openpyxl import load_workbook # Adjust the column width
from openpyxl.drawing.image import Image  # Import for inserting images

# Function to marble dropping and estimate π
def simulate_pi_estimation(total_drops, iterations=10):
    estimated_values = []

    for iteration in range(iterations):
        inside_circle = 0

        # Generate random points and check
        for _ in range(total_drops):
            x_coord = random.uniform(-1, 1)
            y_coord = random.uniform(-1, 1)
            if x_coord**2 + y_coord**2 <= 1:
                inside_circle += 1

        # Calculate π estimate for this iteration
        probability = inside_circle / total_drops
        pi_estimate = 4 * probability
        estimated_values.append(pi_estimate)

    return estimated_values

# Save results to an Excel file with the plot
def export_results_to_excel(data):
    excel_data = {
        "Total Points": [],
        "Mean of π": [],
        "Mode of π": [],
        "All Estimates": []
    }

    for entry in data:
        excel_data["Total Points"].append(entry["Total Drops"])
        excel_data["Mean of π"].append(entry["Mean π"])
        excel_data["Mode of π"].append(entry["Mode π"])
        excel_data["All Estimates"].append(", ".join(f"{val:.6f}" for val in entry["Estimates"]))

    # Write data to Excel
    df = pd.DataFrame(excel_data)
    file_name = "MonteCarlo_Pi_Estimation.xlsx"
    df.to_excel(file_name, index=False)

    # Adjust column widths
    wb = load_workbook(file_name)
    sheet = wb.active

    for col in sheet.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        sheet.column_dimensions[col_letter].width = adjusted_width

    # Insert plot into the Excel sheet
    img_path = "MonteCarlo_Pi_Convergence.png"  
    img = Image(img_path)
    img.anchor = "E2"  
    sheet.add_image(img)

    wb.save(file_name)
    print(f"Results and plot saved to {file_name}")

# Plotting π values against the number of drops
def plot_pi_convergence(data):
    drop_counts = [entry["Total Drops"] for entry in data]
    mean_pi_values = [entry["Mean π"] for entry in data]

    plt.figure(figsize=(10, 6))
    plt.plot(drop_counts, mean_pi_values, marker="o", label="Estimated π")
    plt.axhline(y=3.14159, color="red", linestyle="--", label="True π")
    plt.xscale("log")
    plt.xlabel("Number of Points (Log Scale)")
    plt.ylabel("Estimated π")
    plt.title("Monte Carlo Simulation: π Estimation vs. Total Points")
    plt.legend()
    plt.grid(True)
    plt.savefig("MonteCarlo_Pi_Convergence.png")
    plt.show()

# Main execution
if __name__ == "__main__":
    # Define experiments
    trials = [1000, 10000, 100000, 1000000]
    final_results = []

    print(f"{'Total Points':<15}{'Mean π':<15}{'Mode π':<15}")
    print("-" * 45)

    # Run simulations for each trial count
    for drops in trials:
        pi_estimates = simulate_pi_estimation(drops, iterations=10)
        mean_pi = mean(pi_estimates)
        mode_pi = mode(pi_estimates)

        # Collect results
        final_results.append({
            "Total Drops": drops,
            "Mean π": mean_pi,
            "Mode π": mode_pi,
            "Estimates": pi_estimates
        })

        print(f"{drops:<15}{mean_pi:<15.6f}{mode_pi:<15.6f}")

    # Export results to Excel with the plot
    export_results_to_excel(final_results)

    # Generate plot for convergence
    plot_pi_convergence(final_results)
