#DEPRECATED FIRST VERSION IN CONSOLE
import ordinal

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
# Define the ordinal function
def ordinal(n):
    if 10 <= n % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"

def process_data(all_rows):
    # Extract values from columns
    a_values = [row[0] for row in all_rows]
    b_values = [row[1] for row in all_rows]
    c_values = [row[2] for row in all_rows]
    d_values = [row[3] for row in all_rows]
    e_values = [row[4] for row in all_rows]
    f_values = [row[5] for row in all_rows]
    g_values = [row[6] for row in all_rows]
    h_values = [row[7] for row in all_rows]
    i_values = [row[8] for row in all_rows]
    j_values = [row[9] for row in all_rows]
    k_values = [row[10] for row in all_rows]
    l_values = [row[11] for row in all_rows]
    m_values = [row[12] for row in all_rows]
    n_values = [row[13] for row in all_rows]
    o_values = [row[14] for row in all_rows]
    p_values = [row[15] for row in all_rows]
    q_values = [row[16] for row in all_rows]
    r_values = [row[17] for row in all_rows]
    s_values = [row[18] for row in all_rows]
    t_values = [row[19] for row in all_rows]
    u_values = [row[20] for row in all_rows]
    v_values = [row[21] for row in all_rows]
    w_values = [row[22] for row in all_rows]
    x_values = [row[23] for row in all_rows]
    y_values = [row[24] for row in all_rows]


    # Print or use the extracted values as needed
    print(a_values)
    print(b_values)
    print(c_values)
    print(d_values)
    print(e_values)
    print(f_values)
    print(g_values)
    print(h_values)
    print(i_values)
    print(j_values)
    print(k_values)
    print(l_values)
    print(m_values)
    print(n_values)
    print(o_values)
    print(p_values)
    print(q_values)
    print(r_values)
    print(s_values)
    print(t_values)
    print(u_values)
    print(v_values)
    print(w_values)
    print(x_values)
    print(y_values)



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




if __name__ == "__main__":
    file_path = 'data.xlsx'
    data = load_excel_data(file_path)
    process_data(data)
