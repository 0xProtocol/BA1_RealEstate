import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from streamlit.components.v1 import components, html
from scipy.stats import chisquare

from load_excel import load_excel_data
from collections import Counter
from scipy.stats import kendalltau

# define a class for color codes in the terminal output
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

# configure Streamlit page
st.set_page_config(layout="wide",page_title="Analytics", page_icon="📈")
# add external CSS file for styling
html(
    """
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/boxicons@2.1.0/css/boxicons.min.css">
    """,
    height=0,
)

# list of suggestions for search
suggestions = ["","gender", "optimism", "property type", "age distribution", "place", "interviewer", "housing politics", "rent amount", "housing policy", "chi"]

# suggestion dropdown
search_query = st.selectbox("", suggestions)

# add custom CSS for styling
st.markdown("""
    <style>
        /* Adjust the margin-top value to move the search input higher */
        .st-ax {
            margin-top: -100px !important;
        }
         .st-cu {
        margin-top: 200px;
        /* Adjust the margin-top value as needed */
    }
    </style>
""", unsafe_allow_html=True)

# add custom CSS to center-align the columns
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

# determine the ordinal suffix
def ordinal(n):
    if 10 <= n % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"


with open('style.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

# read data
data = load_excel_data("data.xlsx")

# extract values from columns
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

# get total number of rows and columns in the dataset
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

######### START GUI #########

def custom_card(title, total, from_last_week):
    return f"""  
        <div class="custom-card">
            <h6>{title}</h6>
            <h5>{total}</h5>
            <p>{from_last_week}</p>  
        </div>
    """

def custom_card2(title, total, from_last_week):
    return f"""
        <div class="custom-card2">
            <h6>{title}</h6>
            <h5>{total}</h5>
            <p>{from_last_week}</p>
        </div>
    """

# Create three cards and display them side by side
col1, col2, col3,col4 = st.columns(4)

st.markdown("""
    <style>
        /* Adjust the margin-top value to move the cards higher */
        .custom-card, .custom-card2 {
            margin-top: -55px !important;
        }
    </style>
""", unsafe_allow_html=True)


with col1:
    st.markdown(custom_card("Interviewees", total_rows, ""), unsafe_allow_html=True)

with col2:
    st.markdown(custom_card("Number of key attributes", total_columns, ""), unsafe_allow_html=True)

with col3:
    st.markdown(custom_card2("Place of the statistical survey", "Vienna", ""), unsafe_allow_html=True)

with col4:
    st.markdown(custom_card2("Number of charts", 10, ""), unsafe_allow_html=True)


# data for the gender pie chart
labels1 = ['Men', 'Women', 'None', 'Divers']
values1 = [percentage_men, percentage_women, percentage_none, percentage_divers]
colors1 = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12']
fig1 = go.Figure(
    data=[go.Pie(labels=labels1, values=values1, textinfo='label+percent', hole=0.3, marker=dict(colors=colors1))])
fig1.update_layout(title_text='Gender Distribution', title_font_color='white', title_font_size=23)
fig1.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig1.update_layout(
    title_text='Gender Distribution',
    title_font_color='white',
    title_font_size=23,
    title_x=0.03,
    title_y=0.94,
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white'
)
fig1.update_layout(
    {
        "paper_bgcolor": "rgba(48, 52, 68, 0)",
        "plot_bgcolor": "rgba(0, 0, 0, 0)",
    }
)
fig1.update_traces(textfont_color='white')

# house hunting optimism chart
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
fig2.update_layout(
    title_font_color='white',
    title_font_size=23,
    title_x=0.03,  # Set the title's x position to the center (0 to 1)
    title_y=0.94,  # Set the title's y position (0 to 1)
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white'
)
fig2.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig2.update_layout(
    {
        "paper_bgcolor": "rgba(52,60,76, 0)",
        "plot_bgcolor": "rgba(0, 0, 0, 0)",
    }
)
fig2.update_traces(textfont_color='white')


# property type chart
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
colors2 = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#34495e', '#e67e22']

fig3 = go.Figure(
    data=[go.Bar(x=labels2, y=values2, marker_color=colors2)],
)

fig3.update_layout(title_text='Property type', title_font_color='white', title_font_size=23)
fig3.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig3.update_layout(
    title_font_color='white',
    title_font_size=23,
    title_x=0.03,
    title_y=0.94,
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white'
)
fig3.update_layout(
    {
        "paper_bgcolor": "rgba(48, 52, 68, 0)",
        "plot_bgcolor": "rgba(0, 0, 0, 0)",
    }
)
fig3.update_traces(marker_line_color='black', marker_line_width=1.5, opacity=0.8)


# Distribution of Ages chart
w_values = [age for age in w_values if age is not None]

# Calculate number of bins dynamically
num_bins = int(np.sqrt(len(w_values)))

# Create an array of colors for each bin
colors = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728',
    '#9467bd', '#8c564b', '#e377c2', '#7f7f7f',
    '#bcbd22', '#17becf', '#aec7e8', '#ffbb78',
    '#98df8a', '#ff9896', '#c5b0d5', '#c49c94'
]

# Create a histogram with specified colors for fig4
fig4 = go.Figure(go.Histogram(x=w_values, nbinsx=num_bins, marker_color=colors))

# Update layout
fig4.update_layout(
    title_text='Distribution of Ages',
    title_font_color='white',
    title_font_size=23,
    title_x=0.03,
    title_y=0.94,
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white'
)

fig4.update_layout(
    {
        "paper_bgcolor": "rgba(56,60,76, 0)",
        "plot_bgcolor": "rgba(0, 0, 0, 0)",
    }
)




# place chart
word_counts = Counter(d_values)
most_common_words = word_counts.most_common(10)

labels5 = []
values5 = []

for i, (word, count) in enumerate(most_common_words, 1):
    label = f"{word}"
    value = count
    labels5.append(label)
    values5.append(value)

# Using a different color scale for better visibility
colors5 = ['#1f77b4', '#aec7e8', '#ffbb78', '#2ca02c', '#98df8a', '#d62728', '#ff9896', '#9467bd', '#c5b0d5', '#8c564b']


# Create a horizontal bar chart for fig5
fig5 = go.Figure(go.Bar(x=values5, y=labels5, orientation='h', marker_color=colors5))

# Update layout
fig5.update_layout(
    title_text='Place',
    title_font_color='white',
    title_font_size=23,
    title_x=0.03,
    title_y=0.94,
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white'
)

fig5.update_layout(
    {
        "paper_bgcolor": "rgba(48, 52, 68, 0)",
        "plot_bgcolor": "rgba(0, 0, 0, 0)",
    }
)

fig5.update_traces(textfont_color='white')


#Interviewer chart
word_counts = Counter(e_values)

# Find the most common words and their counts
most_common_words = word_counts.most_common(14)

# Initialize arrays for labels and values
labels2 = []
values2 = []

for i, (word, count) in enumerate(most_common_words, 1):
    label = f"{word}"
    value = count
    labels2.append(label)
    values2.append(value)
colors2 = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12']
fig6 = go.Figure(
    data=[go.Pie(labels=labels2, values=values2, textinfo='label+percent', hole=0.3, marker=dict(colors=colors2))])
fig6.update_layout(title_text='Interviewer', title_font_color='white', title_font_size=23)
fig6.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig6.update_layout(
    title_font_color='white',
    title_font_size=23,
    title_x=0.03,  # Set the title's x position to the center (0 to 1)
    title_y=0.94,  # Set the title's y position (0 to 1)
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white'
)
fig6.update_layout(
    {
        "paper_bgcolor": "rgba(48, 52, 68, 0)",
        "plot_bgcolor": "rgba(0, 0, 0, 0)",
    }
)
fig6.update_traces(textfont_color='white')

# wohnpolitik chart
word_counts = Counter(j_values)
most_common_words = word_counts.most_common(7)
labels2 = []
values2 = []

for i, (word, count) in enumerate(most_common_words, 1):
    label = f"{word}"
    value = count
    labels2.append(label)
    values2.append(value)

colors2 = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12']
fig7 = go.Figure(
    data=[go.Pie(labels=labels2, values=values2, textinfo='label+percent', hole=0.3, marker=dict(colors=colors2))])
fig7.update_layout(title_text='Housing politics', title_font_color='white', title_font_size=23)
fig7.update_layout(
    title_font_color='white',
    title_font_size=23,
    title_x=0.03,  # Set the title's x position to the center (0 to 1)
    title_y=0.94,  # Set the title's y position (0 to 1)
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white'
)
fig7.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig7.update_layout(
    {
        "paper_bgcolor": "rgba(52,60,76, 0)",
        "plot_bgcolor": "rgba(0, 0, 0, 0)",
    }
)
fig7.update_traces(textfont_color='white')

# rent amount chart
word_counts = Counter(h_values)
most_common_words = word_counts.most_common(7)
labels2 = []
values2 = []

for i, (word, count) in enumerate(most_common_words, 1):
    label = f"{word}"
    value = count
    labels2.append(label)
    values2.append(value)

colors2 = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12']
fig8 = go.Figure(
    data=[go.Pie(labels=labels2, values=values2, textinfo='label+percent', hole=0.3, marker=dict(colors=colors2))])
fig8.update_layout(title_text='Rent amount', title_font_color='white', title_font_size=23)
fig8.update_layout(
    title_font_color='white',
    title_font_size=23,
    title_x=0.03,  # Set the title's x position to the center (0 to 1)
    title_y=0.94,  # Set the title's y position (0 to 1)
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white'
)
fig8.update_layout(paper_bgcolor='black', plot_bgcolor='black', font_color='white')
fig8.update_layout(
    {
        "paper_bgcolor": "rgba(52,60,76, 0)",
        "plot_bgcolor": "rgba(0, 0, 0, 0)",
    }
)
fig8.update_traces(textfont_color='white')

# FIGURE -> Housing policy (broken down by gender)
word_counts = Counter(j_values)
most_common_words = word_counts.most_common(7)
adjusted_j_values = [j * 1.059938212 if x == 'M' else j * 0.948675663 if x == 'W' else j for x, j in zip(x_values, j_values) if j is not None]
adjusted_j_values_men = [j * 1.059938212 for x, j in zip(x_values, j_values) if x == 'M' and j is not None]
adjusted_j_values_women = [j * 0.948675663 for x, j in zip(x_values, j_values) if x == 'D' and j is not None]

durchschnitt = sum(adjusted_j_values) / len(adjusted_j_values) if adjusted_j_values else 0
durchschnitt2 = sum(adjusted_j_values_men) / len(adjusted_j_values_men) if adjusted_j_values_men else 0
durchschnitt3 = sum(adjusted_j_values_women) / len(adjusted_j_values_women) if adjusted_j_values_women else 0

labels9 = ['MIXED','MEN','WOMAN']
values9 = [durchschnitt,durchschnitt2,durchschnitt3]
print(Color.MAGENTA + f"Total count of valid : {durchschnitt:.2f}")

colors9 = ['#1f77b4', '#aec7e8', '#ffbb78', '#2ca02c', '#98df8a', '#d62728', '#ff9896', '#9467bd', '#c5b0d5', '#8c564b']


# Create a horizontal bar chart for fig5
fig9 = go.Figure(go.Bar(x=values9, y=labels9, orientation='h', marker_color=colors9))

# Update layout
fig9.update_layout(
    title_text='Place',
    title_font_color='white',
    title_font_size=23,
    title_x=0.03,
    title_y=0.94,
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white'
)

fig9.update_layout(
    {
        "paper_bgcolor": "rgba(48, 52, 68, 0)",
        "plot_bgcolor": "rgba(0, 0, 0, 0)",
    }
)

fig9.update_traces(textfont_color='white')



# FIGURE -> CHI Quadrat
category_counts = Counter(j_values)

observed_frequencies = list(category_counts.values())

unique_categories = len(category_counts)

expected_frequencies = [len(j_values) / unique_categories] * unique_categories

chi2_stat, p_value = chisquare(observed_frequencies, f_exp=expected_frequencies)


categories = list(category_counts.keys())
observed_frequencies = list(category_counts.values())

fig10 = go.Figure()


fig10.add_trace(go.Bar(x=categories, y=observed_frequencies, name='Beobachtet', marker_color='blue'))
fig10.add_trace(go.Bar(x=categories, y=expected_frequencies, name='Erwartet', marker_color='red'))

fig10.update_layout(
    title_text='Comparison of observed and expected frequencies',
    title_x=0.03,
    title_y=0.94,
    title_font_color='white',
    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    font_color='white',
    barmode='group'
)

fig10.update_traces(textfont_color='white')

print(f"Chi-Quadrat-Statistik: {chi2_stat}, p-Wert: {p_value}")

# FIGURE -> KENDALL
filtered_o_values = []
filtered_q_values = []
for o, q in zip(o_values, q_values):
    if o is not None and q is not None:
        filtered_o_values.append(o)
        filtered_q_values.append(q)

if filtered_o_values and filtered_q_values:
    tau, p_value = kendalltau(filtered_o_values, filtered_q_values)

    kendall_values = [tau]
    categories = ['Kendall Tau']

    # fig11 = go.Figure()

    #   fig11.add_trace(go.Bar(x=categories, y=kendall_values, name='Kendall Tau', marker_color='blue'))

    # fig11.update_layout(
    #   title_text='Kendall Tau Koeffizient zwischen o_values und q_values',
    #   title_x=0.03,
    #     title_y=0.94,
    #    title_font_color='white',
    #    paper_bgcolor='rgba(0, 0, 0, 0.1)',
    #    plot_bgcolor='rgba(0, 0, 0, 0.1)',
    #    font_color='white'
    # )

#   fig11.update_traces(textfont_color='white')




#search logic
col1, col2 = st.columns(2)

if search_query.strip() == "":
    # Display all figures if the search bar is empty
    col1.plotly_chart(fig1, use_container_width=True)
    col1.plotly_chart(fig2, use_container_width=True)
    col1.plotly_chart(fig3, use_container_width=True)
    col1.plotly_chart(fig7, use_container_width=True)
    col1.plotly_chart(fig9, use_container_width=True)
    #col1.plotly_chart(fig11, use_container_width=True)
    col2.plotly_chart(fig4, use_container_width=True)
    col2.plotly_chart(fig5, use_container_width=True)
    col2.plotly_chart(fig6, use_container_width=True)
    col2.plotly_chart(fig8, use_container_width=True)
    col2.plotly_chart(fig10, use_container_width=True)

else:
    search_query_lower = search_query.lower()

    if "gender" in search_query_lower:
        col1.plotly_chart(fig1, use_container_width=True)
    elif "optimism" in search_query_lower:
        col1.plotly_chart(fig2, use_container_width=True)
    elif "property type" in search_query_lower:
        col1.plotly_chart(fig3, use_container_width=True)
    elif "age distribution" in search_query_lower:
        col2.plotly_chart(fig4, use_container_width=True)
    elif "place" in search_query_lower:
        col2.plotly_chart(fig5, use_container_width=True)
    elif "interviewer" in search_query_lower:
        col2.plotly_chart(fig6, use_container_width=True)
    elif "housing politics" in search_query_lower:
        col2.plotly_chart(fig7, use_container_width=True)
    elif "rent amount" in search_query_lower:
        col2.plotly_chart(fig8, use_container_width=True)
    elif "housing policy" in search_query_lower:
        col1.plotly_chart(fig9, use_container_width=True)
    elif "chi" in search_query_lower:
        col2.plotly_chart(fig10, use_container_width=True)
    else:
        st.write("No matching figure found. Please refine your search.")


