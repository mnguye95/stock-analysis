import matplotlib.pyplot as plt
import pandas
from datetime import date
import sys

def visual_pcf(ticker):
    all_volume = main_df.groupby(['Symbol', 'Type'])['Volume'].sum().reset_index()
    by_exp = main_df.groupby(['Symbol', 'Type', 'Exp Date'])['Volume'].sum().reset_index()

    volume_html = all_volume.pivot_table('Volume', ['Symbol'], 'Type')
    exp_html = by_exp.pivot_table('Volume', ['Exp Date','Symbol'], 'Type')

    volume_html = volume_html.reset_index(0)
    exp_html = exp_html.reset_index(0)
    exp_html = exp_html.reset_index(0)

    exp_html['PCR'] = exp_html['Put'] / exp_html['Call']
    volume_html['PCR'] = volume_html['Put'] / volume_html['Call']

    new_exp = exp_html[['Symbol', 'Exp Date', 'Call', 'Put', 'PCR']]
    new_vol = volume_html[['Symbol', 'Call', 'Put', 'PCR']]

    new_exp.set_index('Symbol', inplace=True)

    print(new_exp.head(20))

    selected = exp_html[exp_html["Symbol"] == ticker]
    selected = selected.dropna()

    plt.plot(selected['Exp Date'], selected['PCR'], color='blue', marker='o')
    plt.title('{} PCR to Exp Date'.format(ticker))
    plt.xlabel('Exp Date')
    plt.ylabel('Put Call Ratio')
    plt.grid(True)
    plt.axhline(y=1, color='red', linestyle='-')
    plt.show()

# exp_html = exp_html.sort_values(by=['Call', 'PCR'], ascending=[False, True])
# volume_html = volume_html.sort_values(by=['Call', 'PCR'], ascending=[False, True])

# volume_html.to_csv('volume.csv', index=False)
# exp_html.to_csv('exp.csv', index=False)

# writer = pandas.ExcelWriter('options.xlsx', engine='xlsxwriter')

# main_df.to_excel(writer, sheet_name='main', index=False)

# exp_html.to_excel(writer, sheet_name='by_expiration', index=False)

# volume_html.to_excel(writer, sheet_name='by_volume', index=False)

# writer.save()

# volume_html.to_html('volume.html', index=False)
# exp_html.to_html('exp.html', index=False)

if __name__ == "__main__":
    try:
        main_df = pandas.ExcelFile('./results/options_{}.xlsx'.format(date.today().strftime("%b-%d-%Y")), engine='openpyxl')
        main_df = main_df.parse("main")
        visual_pcf(sys.argv[1].upper())
    except FileNotFoundError:
        print("Missing today's option data. Please download today's options trading session first.")




