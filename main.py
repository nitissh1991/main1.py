# import pandas as pd
# import matplotlib.pyplot as plt
# import numpy as np
#
# # Reading the Excel file
# excel_file_path = 'C:\\Nitish\\PSSE_Graph\\input.xlsx'
# df = pd.read_excel(excel_file_path)
#
# # User inputs for chart title and colors
# chart_title = input("Enter the chart title: ")
# parameters_list = ['Voltage (V)', 'Power (P)', 'Reactive Power (Q)', 'Real Power Current (Ip)', 'Reactive Power Current (Iq)', 'Frequency (F)']
#
# # Setting up the plot
# fig, ax1 = plt.subplots(figsize=(10, 8))
# ax1.set_xlabel('Time (seconds)')
# ax1.set_xlim(0, 10)
# ax1.set_xticks(range(11))
# plt.title(chart_title)
#
# # Function to set axis colors
# def set_axis_colors(ax, color):
#     ax.spines['left'].set_color(color)
#     ax.tick_params(axis='y', colors=color)
#     ax.yaxis.label.set_color(color)
#
# # Looping through parameters
# for param in parameters_list:
#     if param in df.columns:
#         color = input(f"Enter the color for {param}: ")
#         min_val = float(input(f"Enter the minimum value for {param} axis: "))
#         max_val = float(input(f"Enter the maximum value for {param} axis: "))
#         increment = float(input(f"Enter the increment value for {param} axis: "))
#         y_ticks = np.arange(min_val, max_val + increment, increment)
#
#         if param == 'Voltage (V)':
#             ax1.set_ylabel(param, color=color)
#             ax1.plot(df['X'], df[param], color=color, label=param)
#             ax1.set_ylim(min_val, max_val)
#             ax1.set_yticks(y_ticks)
#             set_axis_colors(ax1, color)  # Set axis colors
#             # ax1.grid(True)
#             ax1.legend(loc='upper left', bbox_to_anchor=(1.15, 1), borderaxespad=0.)
#
#         elif param == 'Power (P)':
#             ax2 = ax1.twinx()
#             ax2.set_ylabel(param, color=color)
#             ax2.plot(df['X'], df[param], color=color, label=param)
#             ax2.set_ylim(min_val, max_val)
#             ax2.set_yticks(y_ticks)
#             set_axis_colors(ax2, color)  # Set axis colors
#             # ax2.grid(True)
#             ax2.legend(loc='upper left', bbox_to_anchor=(1.15, 0.85), borderaxespad=0.)
#
#         elif param == 'Reactive Power (Q)':
#             ax3 = ax1.twinx()
#             ax3.spines["right"].set_position(("outward", 60))
#             ax3.set_ylabel(param, color=color)
#             ax3.plot(df['X'], df[param], color=color, linestyle='--', label=param)
#             ax3.set_ylim(min_val, max_val)
#             ax3.set_yticks(y_ticks)
#             set_axis_colors(ax3, color)
#             # ax3.grid(True)
#             ax3.legend(loc='upper left', bbox_to_anchor=(1.15, 0.7), borderaxespad=0.)
#
#
# # Otherwise the right y-label is slightly clipped
# fig.tight_layout()
#
# # Show plot with a grid
# ax1.grid(True, which='major', axis='both', color='grey')
#
# # Show the plot
# plt.show()

# import pandas as pd
# import matplotlib.pyplot as plt
# import numpy as np
#
# # Reading the Excel file
# excel_file_path = 'C:\\Nitish\\PSSE_Graph\\input.xlsx'
# excel_file_path2 = 'C:\\Nitish\\PSSE_Graph\\UserInputs.xlsx'
# df_data = pd.read_excel(excel_file_path, sheet_name='Sheet1')  # Replace 'Sheet1' with your data sheet name
# df_inputs = pd.read_excel(excel_file_path2, sheet_name='Sheet1')
#
# # Loop through each row of input data to plot graphs
# for _, row in df_inputs.iterrows():
#     # Extracting user inputs for each graph
#     chart_title = row['ChartTitle']
#     parameters = [row[f'Parameter{i}'] for i in range(1, 4) if row[f'Parameter{i}'] in df_data.columns]
#     colors = [row[f'Color{i}'] for i in range(1, 4) if row[f'Parameter{i}'] in df_data.columns]
#     mins = [row[f'Min{i}'] for i in range(1, 4) if row[f'Parameter{i}'] in df_data.columns]
#     maxes = [row[f'Max{i}'] for i in range(1, 4) if row[f'Parameter{i}'] in df_data.columns]
#     increments = [row[f'Increment{i}'] for i in range(1, 4) if row[f'Parameter{i}'] in df_data.columns]
#
#     # X-axis parameters
#     x_min = row['XMin']
#     x_max = row['XMax']
#     x_increment = row['XIncrement']
#
#     # Setting up the plot
#     fig, ax1 = plt.subplots(figsize=(10, 8))
#     ax1.set_xlabel('Time (seconds)')
#     ax1.set_xlim(x_min, x_max)
#     ax1.set_xticks(np.arange(x_min, x_max + x_increment, x_increment))
#     plt.title(chart_title)
#
#     # Function to set axis colors
#     def set_axis_colors(ax, color):
#         ax.spines['left'].set_color(color)
#         ax.tick_params(axis='y', colors=color)
#         ax.yaxis.label.set_color(color)
#
#     # Plotting each parameter
#     for i, param in enumerate(parameters):
#         color = colors[i]
#         min_val = mins[i]
#         max_val = maxes[i]
#         increment = increments[i]
#         y_ticks = np.arange(min_val, max_val + increment, increment)
#
#         if i == 0:  # Primary y-axis
#             ax1.set_ylabel(param, color=color)
#             ax1.plot(df_data['X'], df_data[param], color=color, label=param)
#             ax1.set_ylim(min_val, max_val)
#             ax1.set_yticks(y_ticks)
#             set_axis_colors(ax1, color)
#             ax1.legend(loc='upper left', bbox_to_anchor=(1.15, 1), borderaxespad=0.)
#
#         elif i == 1:  # Secondary y-axis
#             ax2 = ax1.twinx()
#             ax2.set_ylabel(param, color=color)
#             ax2.plot(df_data['X'], df_data[param], color=color, label=param)
#             ax2.set_ylim(min_val, max_val)
#             ax2.set_yticks(y_ticks)
#             set_axis_colors(ax2, color)
#             ax2.legend(loc='upper left', bbox_to_anchor=(1.15, 0.85), borderaxespad=0.)
#
#         elif i == 2:  # Tertiary y-axis
#             ax3 = ax1.twinx()
#             ax3.spines["right"].set_position(("outward", 60))
#             ax3.set_ylabel(param, color=color)
#             ax3.plot(df_data['X'], df_data[param], color=color, linestyle='--', label=param)
#             ax3.set_ylim(min_val, max_val)
#             ax3.set_yticks(y_ticks)
#             set_axis_colors(ax3, color)
#             ax3.legend(loc='upper left', bbox_to_anchor=(1.15, 0.7), borderaxespad=0.)
#
#     # Otherwise the right y-label is slightly clipped
#     fig.tight_layout()
#
#     #  Show plot with a grid
#     ax1.grid(True, which='major', axis='both', color='grey')
#
#     # Show the plot
#     plt.show()


import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Reading the Excel file
excel_file_path = 'C:\\Nitish\\PSSE_Graph\\case3aFRTriseS2.xlsx'
user_inputs_path = 'C:\\Nitish\\PSSE_Graph\\UserInputs.xlsx'
df_data = pd.read_excel(excel_file_path, sheet_name='Sheet1', skiprows=3)
df_inputs = pd.read_excel(user_inputs_path, sheet_name='Sheet1')

x_column_name = 'Time(s)'

# Loop through each row of input data to plot graphs
for index, row in enumerate(df_inputs.iterrows()):
    _, row = row  # unpacking the tuple returned by iterrows()

    # Extracting user inputs for each graph
    chart_title = row['ChartTitle']
    x_min = row['XMin']
    x_max = row['XMax']
    x_increment = row['XIncrement']

    # Prepare lists for parameters, legends, and other settings
    parameters = []
    legends = []
    colors = []
    mins = []
    maxes = []
    increments = []

    # Loop to process up to 3 parameters and their settings
    for i in range(1, 4):
        seq_num = row.get(f'Sequential{i}', None)  # Sequential number
        legend = row.get(f'Legend{i}', '')  # Corresponding legend

        if seq_num is not None and 0 < seq_num < len(df_data.columns):
            param = df_data.columns[int(seq_num)]  # Convert seq_num to int and get the corresponding column name
            parameters.append(param)
            legends.append(legend if legend else param)  # Use provided legend or column name as default
            colors.append(row[f'Color{i}'])
            mins.append(row[f'Min{i}'])
            maxes.append(row[f'Max{i}'])
            increments.append(row[f'Increment{i}'])

    # Create a new figure for each plot
    fig, ax1 = plt.subplots(figsize=(10, 8))
    ax1.set_xlabel('Time (seconds)')
    ax1.set_xlim(x_min, x_max)
    ax1.set_xticks(np.arange(x_min, x_max + x_increment, x_increment))
    plt.title(chart_title)

    # Function to set axis colors
    def set_axis_colors(ax, color):
        ax.spines['left'].set_color(color)
        ax.tick_params(axis='y', colors=color)
        ax.yaxis.label.set_color(color)

    # Store line objects and their labels for the legend
    lines, labels = [], []

    # Plotting each parameter with legends
    for i, (param, legend) in enumerate(zip(parameters, legends)):
        color = colors[i]
        min_val = mins[i]
        max_val = maxes[i]
        increment = increments[i]
        y_ticks = np.arange(min_val, max_val + increment, increment)

        ax = ax1 if i == 0 else ax1.twinx()  # Primary or secondary y-axis
        line = ax.plot(df_data[x_column_name], df_data[param], color=color, label=legend)

        # Add line and label for legend
        lines.extend(line)
        labels.append(legend)

        if i == 2:  # Tertiary y-axis (if needed)
            ax.spines["right"].set_position(("outward", 60))
        ax.set_ylabel(legend, color=color)
        ax.set_ylim(min_val, max_val)
        ax.set_yticks(y_ticks)
        set_axis_colors(ax, color)

    # Update the legend position to be at the bottom of the plot, arranged horizontally
    ax1.legend(lines, labels, loc='lower center', bbox_to_anchor=(0.5, -0.15), ncol=len(parameters), borderaxespad=0.)

    # Finalizing and Displaying the Plot
    fig.tight_layout()  # Adjust layout
    ax1.grid(True, which='major', axis='both', color='grey')

# Show all plots at once
plt.show()
