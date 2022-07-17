import functools
import streamlit as st
from st_aggrid import AgGrid
from st_aggrid.shared import JsCode
from st_aggrid.grid_options_builder import GridOptionsBuilder
import pandas as pd
import plotly.express as px
from io import BytesIO
from pandas.tseries.offsets import *



chart = functools.partial(st.plotly_chart, use_container_width=True)
COMMON_ARGS = {
    "color": "Total_Routes",
    "color_discrete_sequence": px.colors.sequential.Reds,
    "hover_data": [
        "Total_Routes",
        "Stops",
        "Total_Lines",
    ],
}


@st.experimental_memo
def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.drop(df.index[0])
    df = df.drop(df.index[0])
    df = df.drop(df.index[0])
    df = df.drop(df.index[0])
    df.columns = df.iloc[0]
    df = df.drop(df.index[0])
    df1 = df.copy()
    df1[df1['Customer'].isnull()]=df1[df1['Customer'].isnull()].shift(-1,axis=1)
    df1['Cases'] = pd.to_numeric(df1['Cases'], errors='coerce').fillna(0)
    df1 = df1.dropna(axis=0, subset=['Invoice'])
    df1 = df1.rename(columns={'NO':'Warehouse','DC-WH':'Customer Number','Customer':'Commodity','SCTN':'Invoice','Cases': 'Total_Weight', 'Lines': 'Total_Cases', 'Route': 'Total_Lines', '#': 'Stops', 'Invoice': 'Total_Routes', 'Section': 'Invoice','Weight':'Total_Cube'})
    df1 = df1.drop(columns=['Cube'])
    df1['Stops']=pd.to_numeric(df1['Stops'], downcast='integer', errors='coerce')
    df1 = df1.dropna(axis=0, subset=['Stops'])
    df1['Total_Cases'] = df1['Total_Cases'].astype(float)
    df1['Total_Lines'] = df1['Total_Lines'].astype(float)
    numbers = df1.groupby(['Total_Routes'], as_index=False).sum()
    numbers.reset_index(drop=True, inplace=True)
    stops = df1[['Total_Routes','Stops']]
    stops = stops.groupby('Total_Routes')['Stops'].nunique()
    pivot = numbers.merge(stops, how='inner', left_on=['Total_Routes'], right_on=['Total_Routes'])
    pivot['Total_Routes'] = pivot['Total_Routes'].astype(str)
    pivot = pivot[pivot['Total_Routes'].str.contains('-')==True ]
    pivot = pivot.drop(columns = 'Stops_x')
    pivot = pivot.rename(columns = {'Stops_y':'Stops'})
    return pivot

@st.experimental_memo
def filter_data(
    df: pd.DataFrame, account_selections: list[str], #symbol_selections: list[str#]
) -> pd.DataFrame:
    """
    Returns Dataframe with only accounts and symbols selected
    Args:
        df (pd.DataFrame): clean fidelity csv data, including account_name and symbol columns
        account_selections (list[str]): list of account names to include
        symbol_selections (list[str]): list of symbols to include
    Returns:
        pd.DataFrame: data only for the given accounts and symbols
    """
    df = df.copy()
    df = df[
        df.Total_Routes.isin(account_selections)]
    

    return df


def main() -> None:
    st.header("Team Selection Overview :moneybag: :dollar: :bar_chart:")

    #with st.expander("How to Use This"):
    #    st.write(Path("README.md").read_text())

    st.subheader("Upload your XLSX from EXE")
    #col_names = ["col1",'Order','NO','DC','WH','Customer','SCTN','Invoice','#','Route','Lines','Cases','Weight','Cube']
    #uploaded_data = st.file_uploader(
    #    "Drag and Drop or Click to Upload", type=".xlsx", accept_multiple_files=False)
    uploaded_data = st.file_uploader(
        "Drag and Drop or Click to Upload")
    
   # if uploaded_data is not None:
    #    uploaded_data = pd.read_excel(uploaded_data)
    

    if uploaded_data is None:
        st.info("Upload a file above to use your own data!")
        uploaded_data = '' #pd.read_excel(r'C:\Users\swade\Desktop\test2.xlsx')
    else:
        st.success("Uploaded your file!")
    #col_names = ["col1",'Order','NO','DC','WH','Customer','SCTN','Invoice','#','Route','Lines','Cases','Weight','Cube']
    try:
        df = pd.read_excel(uploaded_data)
        df= df.astype('string')

        with st.expander("Raw Dataframe"):
            st.write(df)

        df = clean_data(df)


        #with st.expander("Cleaned Data"):
        #   st.write(df)

        st.sidebar.subheader("Filter Displayed Routes")

        accounts = list(df.Total_Routes.unique())
        account_selections = st.sidebar.multiselect(
            "Select Routes to View", options=accounts, default=accounts
        )
        #st.sidebar.subheader("Filter Displayed Tickers")

        #symbols = list(df.loc[df.Stops.isin(account_selections), "Stops"].unique())
        #symbol_selections = st.sidebar.multiselect(
        #    "Select Ticker Symbols to View", options=symbols, default=symbols
        #)

        df = filter_data(df, account_selections)
        st.subheader("Selected Route and Stop Data")
        cellsytle_jscode = JsCode(
            """
        function(params) {
            if (params.value > 0) {
                return {
                    'color': 'white',
                    'backgroundColor': 'forestgreen'
                }
            } else if (params.value < 0) {
                return {
                    'color': 'white',
                    'backgroundColor': 'crimson'
                }
            } else {
                return {
                    'color': 'white',
                    'backgroundColor': 'slategray'
                }
            }
        };
        """
        )

        gb = GridOptionsBuilder.from_dataframe(df)
        gb.configure_columns(
            (
                "last_price_change",
                "Total_Lines",
                "total_gain_loss_percent",
                "today's_gain_loss_dollar",
                "today's_gain_loss_percent",
            ),
            #cellStyle=cellsytle_jscode,
        )
        gb.configure_pagination()
        gb.configure_columns(("Total_Routes", "Stops"), pinned=True)
        gridOptions = gb.build()

        AgGrid(df, gridOptions=gridOptions, allow_unsafe_jscode=True)

        def draw_bar(y_val: str) -> None:
            fig = px.bar(df, y=y_val, x="Stops", **COMMON_ARGS)
            fig.update_layout(barmode="stack", xaxis={"categoryorder": "total descending"})
            chart(fig)

        account_plural = "s" if len(account_selections) > 1 else ""
        st.subheader(f"Total Cases t{account_plural}")
        totals = df.groupby("Total_Routes", as_index=False).sum()
        if len(account_selections) > 1:
            st.metric(
                "Total Cases",
                f"{totals.Total_Cases.sum():.2f}",
                #f"{totals.Total_Lines.sum():.2f}",
            )
    #  for column, row in zip(st.columns(len(totals)), totals.itertuples()):
    #      column.metric(
    #          row.Total_Routes,
    #          f"{row.Total_Cases:.2f}",
    #          f"{row.Total_Lines:.2f}",
    #      )

        fig = px.bar(
            totals,
            y="Total_Routes",
            x="Total_Cases",
            color="Total_Routes",
            color_discrete_sequence=px.colors.sequential.Greens,
        )
        fig.update_layout(barmode="stack", xaxis={"categoryorder": "total descending"})
        chart(fig)

        st.subheader("Value of each Stop")
        draw_bar("Total_Cases")

        st.subheader("Value of each Stop per Route")
        fig = px.sunburst(
            df, path=["Total_Routes", "Stops"], values="Total_Cases", **COMMON_ARGS
        )
        fig.update_layout(margin=dict(t=0, b=0, l=0, r=0))
        chart(fig)

        st.subheader("Route Stop Breakdown from Total Routes")
        fig = px.pie(df, values="Total_Cases", names="Stops", **COMMON_ARGS)
        fig.update_layout(margin=dict(t=0, b=0, l=0, r=0))
        chart(fig)
        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, index=True, sheet_name='Most Lines Per Route')
            workbook = writer.book
            #worksheet = writer.sheets['Sheet1']
            format1 = workbook.add_format({'num_format': '0.00'}) 
            #worksheet.set_column('A:A', None, format1)  
            writer.save()
            processed_data = output.getvalue()
            return processed_data
        df_xlsx = to_excel(df)
        st.download_button(label='📥 Download Current Result', data=df_xlsx ,file_name= 'Most_Lines_Per_Route.xlsx')
    except FileNotFoundError:
        st.write("")


if __name__ == "__main__":
    st.set_page_config(
        "EXE Team Selection Work Overview",
        "📊",
        initial_sidebar_state="auto",
        layout="centered",
    )
    main()
