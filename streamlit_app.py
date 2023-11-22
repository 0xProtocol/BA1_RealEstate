import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt


from load_excel import load_excel_data
from collections import Counter
class Color:
    RESET = '\033[0m'
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def ordinal(n):
    if 10 <= n % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"


with open('style.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

data = load_excel_data("data.xlsx")
# Extract values from columns
a_values = [row[0] for row in data]
b_values = [row[1] for row in data]
c_values = [row[2] for row in data]
d_values = [row[3] for row in data]
e_values = [row[4] for row in data]
f_values = [row[5] for row in data]
g_values = [row[6] for row in data]
h_values = [row[7] for row in data]
i_values = [row[8] for row in data]
j_values = [row[9] for row in data]
k_values = [row[10] for row in data]
l_values = [row[11] for row in data]
m_values = [row[12] for row in data]
n_values = [row[13] for row in data]
o_values = [row[14] for row in data]
p_values = [row[15] for row in data]
q_values = [row[16] for row in data]
r_values = [row[17] for row in data]
s_values = [row[18] for row in data]
t_values = [row[19] for row in data]
u_values = [row[20] for row in data]
v_values = [row[21] for row in data]
w_values = [row[22] for row in data]
x_values = [row[23] for row in data]
y_values = [row[24] for row in data]

# Statistics IN WEB APP DISPLAY
print(Color.GREEN + "-------------- STATISTICS GENDER -----------------------")
total_count = len(x_values)
count_men = x_values.count('M')
count_women = x_values.count('W')
count_divers = x_values.count('D')
count_none = x_values.count(None)
percentage_men = (count_men / total_count) * 100
percentage_women = (count_women / total_count) * 100
percentage_divers = (count_divers / total_count) * 100
percentage_none = (count_none / total_count) * 100

# Print the percentages
print(Color.RED + f"Percentage of Men: {percentage_men:.2f}%")
print(Color.RED + f"Percentage of Women: {percentage_women:.2f}%")
print(Color.RED + f"Percentage of Divers: {percentage_divers:.2f}%")
print(Color.RED + f"Percentage of None: {percentage_none:.2f}%")

print(Color.GREEN + "-------------- STATISTICS AGE -----------------------")
non_none_values = [value for value in w_values if value is not None]
total_count_age = len(non_none_values)

sum_age = sum(non_none_values)
average_age = sum_age / total_count_age

print(Color.BLUE + f"Total count of valid age Interviewees: {total_count_age:.2f}")
print(Color.BLUE + f"Average Age of the Interviewees: {average_age:.2f}")

print(Color.GREEN + "-------------- STATISTICS Mietendeckel -----------------------")
non_none_values_u = [value for value in u_values if value is not None]
total_count_mietendeckel = len(non_none_values_u)

sum_mietendeckel = sum(non_none_values_u)
average_mietendeckel = sum_mietendeckel / total_count_mietendeckel

print(Color.YELLOW + f"Total count of valid mietendeckel Interviewees: {total_count_mietendeckel:.2f}")
print(Color.YELLOW + f"Average Mietendeckel of the Interviewees: {average_mietendeckel:.2f}")

print(Color.GREEN + "-------------- STATISTICS Makler harte Arbeit -----------------------")
non_none_values_o = [value for value in o_values if value is not None]
total_count_makler = len(non_none_values_o)

sum_makler = sum(non_none_values_o)
average_makler = sum_makler / total_count_makler

print(Color.CYAN + f"Total count of valid makler Interviewees: {total_count_makler:.2f}")
print(Color.CYAN + f"Average makler hard job of the Interviewees: {average_makler:.2f}")

print(Color.GREEN + "-------------- STATISTICS Wohnungssuche Optimismus -----------------------")
non_none_values_i = [value for value in i_values if value is not None]
total_count_optimismus = len(non_none_values_i)

sum_optimismus = sum(non_none_values_i)
average_optimismus = sum_optimismus / total_count_optimismus

print(Color.MAGENTA + f"Total count of valid makler Interviewees: {total_count_optimismus:.2f}")
print(Color.MAGENTA + f"Average optimism rentsearch of the Interviewees: {average_optimismus:.2f}")

print(Color.GREEN + "-------------- STATISTICS häufigste Art der Wohnung -----------------------")

# Use Counter to count occurrences of each word
word_counts = Counter(g_values)

# Find the most common words and their counts
most_common_words = word_counts.most_common(7)

for i, (word, count) in enumerate(most_common_words, 1):
    print(Color.RED + f"{i}. The {ordinal(i)} most common word is '{word}' with {count} occurrences.")


#START GUI

##HEADER
st.markdown("<h1 style='text-align: center;'>REAL ESTATE ANALYTICS</h1>", unsafe_allow_html=True)

st.markdown('### Metrics')
col1, col2, col3, col4 = st.columns(4)
col1.metric("Temperature", percentage_men, percentage_women)
col2.metric("Wind", "9 mph", "-8%")
col3.metric("Humidity", "86%", "4%")
col4.metric("Humidity", "86%", "4%")






chart_data = pd.DataFrame(np.random.randn(20, 3), columns=["a", "b", "c"])
st.bar_chart(chart_data)


# Pie chart, where the slices will be ordered and plotted counter-clockwise:
labels = 'Men', 'Women ', 'None', 'Divers'
sizes = [percentage_men, percentage_women, percentage_none, percentage_divers]
explode = (0, 0.1, 0, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')

fig1, ax1 = plt.subplots()
ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
        shadow=True, startangle=90)
ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

st.pyplot(fig1)


##
# Data
categories = ['Category 1', 'Category 2', 'Category 3', 'Category 4']
variable1 = [10, 15, 20, 25]
variable2 = [5, 12, 18, 22]
variable3 = [8, 10, 15, 20]
variable4 = [12, 18, 25, 30]

# Set up positions for bars on X-axis
x = np.arange(len(categories))

# Set up the figure and axes
fig, ax = plt.subplots()

# Plot each variable as a separate set of bars
bar_width = 0.2  # Adjust this value based on your preference
ax.bar(x - 1.5 * bar_width, variable1, width=bar_width, label='Variable 1')
ax.bar(x - 0.5 * bar_width, variable2, width=bar_width, label='Variable 2')
ax.bar(x + 0.5 * bar_width, variable3, width=bar_width, label='Variable 3')
ax.bar(x + 1.5 * bar_width, variable4, width=bar_width, label='Variable 4')

# Set up labels and title
ax.set_xticks(x)
ax.set_xticklabels(categories)
ax.set_ylabel('Values')
ax.set_title('Bar Chart with Four Variables')

# Add a legend
ax.legend()

# Save the plot as an image file (e.g., PNG)
plt.savefig('bar_chart.png')

# Alternatively, you can save as a PDF
# plt.savefig('bar_chart.pdf')

# Alternatively, you can display the plot in non-interactive environments
# plt.show()