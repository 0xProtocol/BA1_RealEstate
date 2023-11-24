import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go

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


st.set_page_config(layout="wide",page_title="Analytics", page_icon="📈")

# Add custom CSS to center-align the columns
st.markdown(
    """
    <style>
        .element-container {
            display: flex;
            flex-direction: row;
            justify-content: center;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.sidebar.header("Plotting Demo2")


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

total_rows = len(a_values)
df = pd.read_excel("data.xlsx")
total_columns = len(df.columns)

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

print(Color.GREEN + "-------------- STATISTICS AGE2 -----------------------")
# Filter out None values
non_none_values = [value for value in w_values if value is not None]
# Count occurrences of each age
age_counts = Counter(non_none_values)
# Total count of valid ages
total_count_age = len(non_none_values)
for age, count in age_counts.items():
    percentage = (count / total_count_age) * 100
    print(Color.BLUE + f"Age: {age}, Percentage: {percentage:.2f}%")



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

# Calculate percentages
percentages = [(value / total_count_optimismus) * 100 for value in non_none_values_i]

# Create a Counter to store counts and percentages for each unique value
value_counter = Counter(zip(non_none_values_i, percentages))

# Sum percentages for each unique value
sum_percentages_by_value = {}
for (value, percentage), count in value_counter.items():
    if value not in sum_percentages_by_value:
        sum_percentages_by_value[value] = 0
    sum_percentages_by_value[value] += percentage * count

# Print sum of percentages for each unique value
for value, sum_percentage in sum_percentages_by_value.items():
    print(f"Sum of percentages for value {value}: {sum_percentage:.2f}%")


print(Color.GREEN + "-------------- STATISTICS häufigste Art der Wohnung -----------------------")

# Use Counter to count occurrences of each word
word_counts = Counter(g_values)

# Find the most common words and their counts
most_common_words = word_counts.most_common(7)

for i, (word, count) in enumerate(most_common_words, 1):
    print(Color.RED + f"{i}. The {ordinal(i)} most common word is '{word}' with {count} occurrences.")

# START GUI

##HEADER
st.markdown("<h1 style='text-align: center; pointer-events: none;'>REAL ESTATE ANALYTICS</h1>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)  # Add a line break for space
st.markdown("<br>", unsafe_allow_html=True)  # Add a line break for space
st.markdown("<h2 style='text-align: left; pointer-events: none; color: #33ccff;'>Key Metrics</h2>",
            unsafe_allow_html=True)

# Display metrics side by side

col1, col2, col3 = st.columns(3)

with col1:
    st.metric(label="Number of interviewees", value=total_rows)

with col2:
    st.metric(label="Number of key attributes", value=total_columns)

with col3:
    st.metric(label="Place of the statistical survey", value='Vienna')

st.markdown("<h2 style='text-align: left; color: #33ccff; pointer-events: none;'>Basic Statistics</h2>", unsafe_allow_html=True)
# Example data for the first pie chart
labels1 = ['Men', 'Women', 'None', 'Divers']
values1 = [25, 30, 15, 30]
colors1 = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12']
fig1 = go.Figure(
    data=[go.Pie(labels=labels1, values=values1, textinfo='label+percent', hole=0.3, marker=dict(colors=colors1))])
fig1.update_layout(title_text='Gender Distribution', title_font_color='white', title_font_size=23)
fig1.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig1.update_traces(textfont_color='white')

word_counts = Counter(i_values)
most_common_words = word_counts.most_common(7)
labels2 = []
values2 = []

for i, (word, count) in enumerate(most_common_words, 1):
    label = f"{word}"
    value = count
    labels2.append(label)
    values2.append(value)

colors2 = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12']
fig2 = go.Figure(
    data=[go.Pie(labels=labels2, values=values2, textinfo='label+percent', hole=0.3, marker=dict(colors=colors2))])
fig2.update_layout(title_text='House hunting optimism', title_font_color='white', title_font_size=23)
fig2.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig2.update_traces(textfont_color='white')


# Example data for the second pie chart
# Use Counter to count occurrences of each word
word_counts = Counter(g_values)

# Find the most common words and their counts
most_common_words = word_counts.most_common(7)

# Initialize arrays for labels and values
labels2 = []
values2 = []

for i, (word, count) in enumerate(most_common_words, 1):
    label = f"{word}"
    value = count
    labels2.append(label)
    values2.append(value)
colors2 = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12']
fig3 = go.Figure(
    data=[go.Pie(labels=labels2, values=values2, textinfo='label+percent', hole=0.3, marker=dict(colors=colors2))])
fig3.update_layout(title_text='Distribution of Categories', title_font_color='white', title_font_size=23)
fig3.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig3.update_traces(textfont_color='white')


########################4##############
# Example data for the second pie chart
# Use Counter to count occurrences of each word
# Group ages into ranges of 9 years
# Filter out None values if present
w_values = [age for age in w_values if age is not None]

# Group ages into ranges of 9 years
age_ranges = [(i, i + 9) for i in range(0, max(w_values, default=0) + 1, 9)]
grouped_values = Counter()

for age in w_values:
    for age_range in age_ranges:
        if age_range[0] <= age < age_range[1]:
            grouped_values[age_range] += 1
            break

# Find the most common age ranges and their counts
most_common_age_ranges = grouped_values.most_common(20)

# Initialize arrays for labels and values
labels2 = []
values2 = []

for i, (age_range, count) in enumerate(most_common_age_ranges, 1):
    label = f"{age_range[0]}-{age_range[1]}"
    value = count
    labels2.append(label)
    values2.append(value)

colors2 = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12']

fig4 = go.Figure(
    data=[go.Pie(labels=labels2, values=values2, textinfo='label+percent', hole=0.3, marker=dict(colors=colors2))])

fig3.update_layout(title_text='Most common Object', title_font_color='white', title_font_size=23)
fig4.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig4.update_traces(textfont_color='white')


# Create a 4-column layout
col1, col2 = st.columns(2)

# Display the first pie chart in the first column
with col1:
    st.plotly_chart(fig1, use_container_width=True)

# Display the second pie chart in the second column
with col2:
    st.plotly_chart(fig2, use_container_width=True)

# Create a new column for the third pie chart
col3, col4 = st.columns(2)
with col3:
    st.plotly_chart(fig3, use_container_width=True)
with col4:
    st.plotly_chart(fig4, use_container_width=True)

st.markdown("<h2 style='text-align: left; color: #33ccff; pointer-events: none;'>Advanced Statistics</h2>", unsafe_allow_html=True)


####SIDEBAR####


#https://blog.streamlit.io/introducing-theming/