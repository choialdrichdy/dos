import pandas as pd
import numpy as np
from datetime import date
from config import DISTRIBUTOR_WATCHLIST, MODELS_WATCHLIST, LOCATIONS, INPUT_FILE, OUTPUT_LOCATION


def filter_demo_models(initial_models):
    models = []

    for model in initial_models:
        if 'demo' not in model.lower():
            models.append(model)

    return models


def filter_distributors(initial_distributors):
    distributors = []
    if not DISTRIBUTOR_WATCHLIST:
        return initial_distributors

    for distributor in initial_distributors:
        for watchlist in DISTRIBUTOR_WATCHLIST:
            if watchlist.lower() in distributor.lower():
                distributors.append(distributor)
    return distributors


def generate_dos_template(filtered_df, value_to_use):
    dos_table = pd.pivot_table(filtered_df, values=value_to_use, index=['Territory', 'Statistics Model'],
                        columns=['Distributor'], aggfunc=np.sum, fill_value=0)
    dos_table = dos_table.reset_index()
    dos_table.index.name = None

    d = {'Statistics Model': MODELS_WATCHLIST}
    watchlist_df = pd.DataFrame(data=d)
    for territory in LOCATIONS:
        watchlist_df['Territory'] = territory
        dos_table = pd.merge(dos_table, watchlist_df, how='outer', on=['Territory', 'Statistics Model'])
    
    dos_table = dos_table.sort_values(["Territory", "Statistics Model"])
    dos_table = dos_table.replace(np.nan, 0)
    return dos_table


def generate_dos_file(filtered_df, date):
    filtered_df.loc[~filtered_df["Statistics Model"].isin(MODELS_WATCHLIST), "Statistics Model"] = "Other Models"

    inventory_dos = generate_dos_template(filtered_df, 'Distributor Stock')
    sales_dos = generate_dos_template(filtered_df, 'Sales Volume')

    writer = pd.ExcelWriter(f'{OUTPUT_LOCATION}/National-DOS-{date}.xlsx', engine='openpyxl')
    inventory_dos.to_excel(writer, sheet_name='Inventory', index=False)
    sales_dos.to_excel(writer, sheet_name='Sales', index=False)
    writer.save()


def generate_distributor_file(filtered_df, date):
    d = {'Statistics Model': MODELS_WATCHLIST}
    watchlist_df = pd.DataFrame(data=d)
    for key, value in filtered_df.groupby('Distributor'):
        writer = pd.ExcelWriter(f'{OUTPUT_LOCATION}/{key.lower().replace(" ", "-")}-{date}.xlsx', engine='openpyxl')
        for territory, territory_df in value.groupby('Territory'):
            territory_pivot = pd.pivot_table(territory_df, values='Distributor Stock', index=['Statistics Model'], columns=['Store/Warehouse'], aggfunc=np.sum, fill_value=0).reset_index()
            territory_pivot = pd.merge(territory_pivot, watchlist_df, how='outer', on=['Statistics Model'])
            territory_pivot = territory_pivot.sort_values(["Statistics Model"])
            territory_pivot = territory_pivot.replace(0, np.nan)
            stores = territory_pivot.keys().tolist()
            total = territory_pivot[stores].sum()
            total[0] = 'Total'
            territory_pivot = territory_pivot.append(pd.DataFrame([total], columns=total.index, index=pd.Index([-1]))).sort_index()
            territory_pivot.to_excel(writer, sheet_name=territory, index=False)
        writer.save()


def get_raw_data():
    """
        Get initial raw data from the file location in config.py.
        Will also do some initial preprocessing.
        (remove last row and make negative data to zero)
    """
    df = pd.read_excel(f'{INPUT_FILE}')
    df = df[:-1]
    num = df._get_numeric_data()
    num[num < 0] = 0
    df = df.fillna(0)
    return df


def main():
    today = date.today().strftime("%b-%d-%Y")

    df = get_raw_data()
    distributors = filter_distributors(df.Distributor.unique().tolist())
    models = filter_demo_models(df['Statistics Model'].unique().tolist())

    filtered_df = df[(df['Distributor'].isin(distributors)) & (df['Statistics Model'].isin(models))]
    generate_dos_file(filtered_df.copy(deep=True), today)
    generate_distributor_file(filtered_df.copy(deep=True), today)


if __name__ == "__main__":
    main()