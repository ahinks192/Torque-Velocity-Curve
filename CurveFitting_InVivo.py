import os
import numpy as np
import pandas as pd
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt
from sklearn.metrics import r2_score
from tkinter import Tk
from tkinter.simpledialog import askstring
from tkinter.filedialog import askopenfilename
from matplotlib.offsetbox import AnchoredText

final_results = pd.DataFrame(
    columns=[
        'Animal',
        'F max',
        'V max',
        'Peak power',
        'Optimal Force',
        'Optimal Velocity',
        'a coeff',
        'b coeff',
        'Curvature',
        'R-squared'
    ]
)

# Define the Hill's equation function
def hill_equation(F, F_max, a, b):
    return (((F_max + a)*b)/(F + a)) - b

# Define a new function that fixes F_max and only takes a and b as parameters
def hill_equation_fixed_Fmax(F, a, b):
    F_max = np.max(F)
    return hill_equation(F, F_max, a, b)

# Create a Tinter root winkdow to hide
root = Tk()
root.withdraw()

# Get the file path and sheet name from the user
file_path = askopenfilename(title="Select Excel file")
excel_file = pd.ExcelFile(file_path)
animal_datasheets = []
for sheet in excel_file.sheet_names:
    if 'ISOK' in sheet or 'ISOT' in sheet:
        animal_datasheets.append(sheet)

for idx, sheet in enumerate(animal_datasheets):
    final_results.loc[idx, 'Animal'] = sheet
    data = pd.read_excel(file_path, sheet_name=sheet)

    # Extract the force and velocity data from the Excel sheet
    force = data['force'].values
    velocity = data['velocity'].values

    a_guess = 1
    b_guess = 1

    # Fit the data to Hill's equation using the Trust Region Reflective algorithm
    popt, pcov = curve_fit(hill_equation_fixed_Fmax, force, velocity, p0=[a_guess, b_guess], method='trf')

    # Predict the velocity values using the fitted curve
    velocity_pred = hill_equation(force, np.max(force), *popt)

    # Compute the R-squared value
    r_squared = r2_score(velocity, velocity_pred)
    final_results.loc[idx, 'R-squared'] = float(f'{r_squared:.3f}')
    # Calculate the maximum force
    Fmax = max(force)
    final_results.loc[idx, 'F max'] = float(f'{Fmax:.3f}')

    # Compute the curvature
    curvature = popt[0] / Fmax
    final_results.loc[idx, 'Curvature'] = float(f'{curvature:.3f}')

    # Define the range of forces for the extended curve
    force_extended = np.linspace(0, Fmax, 100)
    # Predict the velocity values for the extended curve
    velocity_extended = hill_equation(force_extended, np.max(force), *popt)
    velocity_extended[force_extended > np.max(force)] = 0

    # Calculate the maximum velocity
    Vmax = max(velocity_extended)
    final_results.loc[idx, 'V max'] = float(f'{Vmax:.3f}')

    # Compute power
    power: pd.DataFrame = pd.DataFrame(data = ((force_extended * velocity_extended)), columns=['Power'])

    PP_index = power.idxmax()
    #Calculate peak power
    peak_power = power['Power'][PP_index].values[0]
    optimal_force = force_extended[PP_index][0]
    optimal_velocity = velocity_extended[PP_index][0]


    power_percents = {}
    for percentage in [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80]:
        val = Fmax * percentage
        power_percents[force_extended[np.argmin(np.abs(force_extended - val))]] = power['Power'][np.argmin(np.abs(force_extended - val))]
        final_results.loc[idx, f'{percentage * 100:.0f}% Power'] = float(f"{power['Power'][np.argmin(np.abs(force_extended - val))]:.3f}")

    # Show the plot
    # Plot the data and the fitted curve
    fig, ax = plt.subplots(figsize=(9,6))
    ax.scatter(force, velocity, label='Data')
    ax.plot(force_extended, velocity_extended, label='Fitted Curve')
    ax.set_xlabel('Force (mN*m)')
    ax.set_ylabel('Velocity (deg/s)', color=(0, 0.447, 0.698))
    ax.axhline(y=optimal_velocity, color='red', linestyle='--', label='Optimal Velocity')
    ax.axvline(x=optimal_force, color='red', linestyle='--', label='Optimal Force')

    ax2=ax.twinx()
    ax2.plot(force_extended, power, label='Power', color='green')
    ax2.plot(optimal_force, peak_power, color = 'red', marker='+', markersize=6)
    ax2.plot(power_percents.keys(), power_percents.values(), color = 'red', marker='+', markersize=12, linestyle='')
    ax2.set_ylabel("Power(mN*m*deg/s)", color='green')
    ax.set_title(f'{sheet}\nF_max={Fmax:.2f}, a={popt[0]:.2f}, b={popt[1]:.2f}, R-squared={r_squared:.2f}')
    ax.legend()

    final_results.loc[idx, 'a coeff'] = float(f'{popt[0]:.3f}')
    final_results.loc[idx, 'b coeff'] = float(f'{popt[1]:.3f}')
    final_results.loc[idx, 'Optimal Force'] = float(f'{optimal_force:.3f}')
    final_results.loc[idx, 'Optimal Velocity'] = float(f'{optimal_velocity:.3f}')
    final_results.loc[idx, 'Peak power'] = float(f'{peak_power:.3f}')

    offset = 0
    for string, value in zip(['Vmax', 'Fmax', 'Curvature', 'PP', 'Opt Force', 'Opt Vel'], [Vmax, Fmax, curvature, peak_power, optimal_force, optimal_velocity]):
        plt.text(0.5, 0.9 - offset, f"{string}={value:.2f}", ha="center", transform=ax.transAxes)
        offset += 0.1

    plt.show()

output_filepath = '/Users/averyhinks/Desktop/excel_output.xlsx'
final_results.to_excel(output_filepath, index=False)
os.system(f'open {output_filepath}')