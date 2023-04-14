import pandas as pd
# excel_file = 'Keval.xlsx'
def run(path):
    #excel_file = 'F:\Visual Studio\smallData.xlsx'
    # excel_file = path
    # df = pd.read_excel(excel_file)
    df = path
    pd.set_option('display.max_columns',105)

    # Convert Order Date to DateTime and Set as Index
    df['Order Date'] = pd.to_datetime(df['Order Date'],format='%B %d, %Y %I:%M:%S %p')

    # Get Final Status
    status = []
    for row, rowInfo in zip(df['Order Status'], df['ExtraInfo'] ):
        if row == "Cancel Init":
            status.append('SHIPPED')
        elif row == "Cancelled":
            status.append('CANCEL')
        elif row == "Cancelled Before Shipping":
            status.append('CANCEL')
        elif row == "Cancelled Return Received":
            status.append('RTO')
        elif row == "Delivered":
            status.append('DELIVERED')
        elif row == "In Transit":
            status.append('SHIPPED')
        elif row == "Partial Cancelled Return Received":
            status.append('DTO')
        elif row == "Partial Return Received":
            status.append('DTO')
        elif row == "Return Init":
            status.append('DTO')
        elif row == "Return Received":
            status.append('DTO')
        elif row == "Shipped":
            if "EXCHANGE" in str(rowInfo):
                status.append('RTO')
                
            else:
                status.append('SHIPPED')

    df['Status'] = status

    # Profit Calculations
    df.loc[(df['Order Status'] == 'Delivered') | (df['Order Status'] == 'Cancel Init') | (
                df['Order Status'] == 'In Transit') | (df['Order Status'] == 'Shipped') | (
                    df['Order Status'] == 'Return Received') | (df['Order Status'] == 'Return Init') | (
                    df['Order Status'] == 'Partial Return Received') | (
                    df['Order Status'] == 'Partial Cancelled Return Received') | (
                    df['Order Status'] == 'Cancelled Return Received'), 'Profit'] = df['Settlement Amount'] + (
                df['Cost Price'] * df['Return Good Qty']) - (df['Cost Price'] * df['Qty'])
    df.loc[(df['Order Status'] == 'Delivered') | (df['Order Status'] == 'Cancel Init') | (
                df['Order Status'] == 'In Transit') | (df['Order Status'] == 'Shipped') | (
                    df['Order Status'] == 'Return Received') | (df['Order Status'] == 'Return Init') | (
                    df['Order Status'] == 'Partial Return Received') | (
                    df['Order Status'] == 'Partial Cancelled Return Received') | (
                    df['Order Status'] == 'Cancelled Return Received'), 'Profit'] = df['Settlement Amount'] + (
                df['Cost Price'] * df['Return Good Qty']) - (df['Cost Price'] * df['Qty'])

    df.loc[(df['Order Status'] == 'Cancelled') | (df['Order Status'] == 'Cancelled Before Shipping'), 'Profit'] = df[
        'Settlement Amount']

    df.loc[(df['Order Status'] == 'Shipped') & ("EXCHANGE" in df['ExtraInfo']), 'Profit'] = df[
        'Settlement Amount']
    # All Groups
    day_wise_grp = df.groupby('Order Date').agg(total_profit= ('Profit' , 'sum'))
    channel_grp = df.groupby('Channel Name')
    channel_product_grp = df.groupby(['Channel Name','Product Name'])
    product_grp = df.groupby(['Product Name'])
    shipping_company_grp = df.groupby(['Shipping Company'])

    # Channel Wise Group Report
    final_total=channel_grp['Status'].count()
    semi_total = channel_grp['Status'].count() - channel_grp['Status'].apply(lambda x:x.str.contains('CANCEL').sum())

    status_delivered = channel_grp['Status'].apply(lambda x:x.str.contains('DELIVERED').sum())
    status_rto = channel_grp['Status'].apply(lambda x:x.str.contains('RTO').sum())
    status_dto = channel_grp['Status'].apply(lambda x:x.str.contains('DTO').sum())
    status_shipped = channel_grp['Status'].apply(lambda x:x.str.contains('SHIPPED').sum())
    status_cancel = channel_grp['Status'].apply(lambda x:x.str.contains('CANCEL').sum())
    # totalQty = channel_grp['Qty'].sum()
    # totalQty = channel_grp.loc[(df['Status'] == 'DELIVERED'), 'Qty'].sum()
    totalQty = channel_grp.apply(lambda x: x[x['Status'] == 'DELIVERED']['Qty'].sum())
    totalProfitByQty = channel_grp.apply(lambda x: x[x['Status'] == 'DELIVERED']['Profit'].sum())
    total_profit = channel_grp['Profit'].sum()
    print("total_profit",total_profit)
    # avgProfit = totalProfitByQty/tota
    avgProfit = totalProfitByQty.div(totalQty.values)
    notDelieveredQty = channel_grp.apply(lambda x: x[x['Status'] != 'DELIVERED']['Qty'].sum())
    notDeliveredProfit = total_profit - totalProfitByQty
    avgNotDelievredProfit = notDeliveredProfit.div(notDelieveredQty.values)
    tf=df['Profit'].sum()

    comboStatus_df = pd.concat([semi_total,status_delivered,status_rto,status_dto,status_shipped,status_cancel,final_total, total_profit, totalQty, totalProfitByQty, avgProfit, notDelieveredQty, notDeliveredProfit, avgNotDelievredProfit],axis='columns',sort=False)
    comboStatus_df.set_axis(["SemiTotal", "DeliveredTotal", 'RtoTotal', 'DtoTotal', 'ShippedTotal', 'CancelTotal', 'TotalOrders','TotalProfit', 'totalQty','totalProfitByQty', 'avgProfit', 'notDelieveredQty', 'notDeliveredProfit', 'avgNotDelievredProfit'], axis="columns",inplace=True)

    # Percentage Calculation
    comboStatus_df['PctProfit']=(comboStatus_df['TotalProfit']/tf)*100
    comboStatus_df['PctDelivered']=(comboStatus_df['DeliveredTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctRto']=(comboStatus_df['RtoTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctDto']=(comboStatus_df['DtoTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctShipped']=(comboStatus_df['ShippedTotal']/comboStatus_df['SemiTotal'])*100

    comboStatus_df['PctProfit'] = comboStatus_df['PctProfit'].round(2)
    comboStatus_df['PctDelivered'] = comboStatus_df['PctDelivered'].round(2)
    comboStatus_df['PctRto'] = comboStatus_df['PctRto'].round(2)
    comboStatus_df['PctDto'] = comboStatus_df['PctDto'].round(2)
    comboStatus_df['PctShipped'] = comboStatus_df['PctShipped'].round(2)
    
    comboStatus_df['PctProfit'] = comboStatus_df['PctProfit'].astype(str) + " %"
    comboStatus_df['PctDelivered'] = comboStatus_df['PctDelivered'].astype(str) + " %"
    comboStatus_df['PctRto'] = comboStatus_df['PctRto'].astype(str) + " %"
    comboStatus_df['PctDto'] = comboStatus_df['PctDto'].astype(str) + " %"
    comboStatus_df['PctShipped'] = comboStatus_df['PctShipped'].astype(str) + " %"

    ChannelWiseReport = comboStatus_df.loc[:,['TotalProfit','PctProfit','TotalOrders','DeliveredTotal','PctDelivered', 'RtoTotal', 'PctRto','DtoTotal','PctDto','ShippedTotal','PctShipped','CancelTotal' ]]
    ChannelWiseReport = ChannelWiseReport.sort_values(by=['TotalProfit'])
    ChannelWiseReport2 = comboStatus_df.loc[:,['totalQty','totalProfitByQty', 'avgProfit', 'notDelieveredQty', 'notDeliveredProfit', 'avgNotDelievredProfit']]
    ChannelWiseReport2['avgProfit'] = ChannelWiseReport2['avgProfit'].fillna(0)
    ChannelWiseReport2['avgProfit'] = ChannelWiseReport2['avgProfit'].round(2)
    ChannelWiseReport2['avgNotDelievredProfit'] = ChannelWiseReport2['avgNotDelievredProfit'].round(2)
    ChannelWiseReport2 = ChannelWiseReport2.sort_values(by=['avgProfit'])


    # print("Testtest", df['Qty'])
    # ChannelWiseReport2 = ChannelWiseReport2.sort_values(by=['TotalProfit'])
    # ChannelWiseReport

    # Channel-Product Wise Group Report
    final_total=channel_product_grp['Status'].count()
    semi_total = channel_product_grp['Status'].count() - channel_product_grp['Status'].apply(lambda x:x.str.contains('CANCEL').sum())

    status_delivered = channel_product_grp['Status'].apply(lambda x:x.str.contains('DELIVERED').sum())
    status_rto = channel_product_grp['Status'].apply(lambda x:x.str.contains('RTO').sum())
    status_dto = channel_product_grp['Status'].apply(lambda x:x.str.contains('DTO').sum())
    status_shipped = channel_product_grp['Status'].apply(lambda x:x.str.contains('SHIPPED').sum())
    status_cancel = channel_product_grp['Status'].apply(lambda x:x.str.contains('CANCEL').sum())

    total_profit = channel_product_grp['Profit'].sum()
    tf=channel_grp['Profit'].sum()

    comboStatus_df = pd.concat([semi_total,status_delivered,status_rto,status_dto,status_shipped,status_cancel,final_total, total_profit],axis='columns',sort=False)
    comboStatus_df.set_axis(["SemiTotal", "DeliveredTotal", 'RtoTotal', 'DtoTotal', 'ShippedTotal', 'CancelTotal', 'TotalOrders','TotalProfit'], axis="columns",inplace=True)

    # Percentage Calculation
    comboStatus_df['PctProfit']=(comboStatus_df['TotalProfit']/tf)*100
    comboStatus_df['PctDelivered']=(comboStatus_df['DeliveredTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctRto']=(comboStatus_df['RtoTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctDto']=(comboStatus_df['DtoTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctShipped']=(comboStatus_df['ShippedTotal']/comboStatus_df['SemiTotal'])*100

    comboStatus_df['PctProfit'] = comboStatus_df['PctProfit'].round(2)
    comboStatus_df['PctDelivered'] = comboStatus_df['PctDelivered'].round(2)
    comboStatus_df['PctRto'] = comboStatus_df['PctRto'].round(2)
    comboStatus_df['PctDto'] = comboStatus_df['PctDto'].round(2)
    comboStatus_df['PctShipped'] = comboStatus_df['PctShipped'].round(2)
    
    comboStatus_df['PctProfit'] = comboStatus_df['PctProfit'].astype(str) + " %"
    comboStatus_df['PctDelivered'] = comboStatus_df['PctDelivered'].astype(str) + " %"
    comboStatus_df['PctRto'] = comboStatus_df['PctRto'].astype(str) + " %"
    comboStatus_df['PctDto'] = comboStatus_df['PctDto'].astype(str) + " %"
    comboStatus_df['PctShipped'] = comboStatus_df['PctShipped'].astype(str) + " %"

    Channel_Product_WiseReport = comboStatus_df.loc[:,['TotalProfit','PctProfit','TotalOrders','DeliveredTotal','PctDelivered', 'RtoTotal', 'PctRto','DtoTotal','PctDto','ShippedTotal','PctShipped','CancelTotal' ]]
    #Channel_Product_Wise Report in Diffrent Sheets
    # print("1Cols r", Channel_Product_WiseReport.columns)
    # Channel_Product_WiseReport =  Channel_Product_WiseReport.sort_values(by=['TotalProfit'])
    # Channel_Product_WiseReport =  Channel_Product_WiseReport.sort_values(by=['Channel Name'])
    # Channel_Product_WiseReport = Channel_Product_WiseReport.groupby(['Channel Name','Product Name'])
    Channel_Product_WiseReport = Channel_Product_WiseReport.groupby(['Channel Name']).apply(lambda x: x.sort_values('TotalProfit'))
    # print("2Cols r", Channel_Product_WiseReport.columns)
    # Channel_Product_WiseReport = Channel_Product_WiseReport.set_index('Channel Name', inplace=False)
    #Channel_Product_WiseReport = Channel_Product_WiseReport.reset_index(drop=True)
    # Channel_Product_WiseReport = Channel_Product_WiseReport.drop('Channel Name')
    # df['Channel Name'].unique()
    # for state in df['Channel Name'].unique():
    #     print(state)

    #     #Blank Template
    #     writer = pd.ExcelWriter("F:\k\Output101.xlsx",engine = 'xlsxwriter')
    #     for state in df['Channel Name'].unique():
    #         NewDf = df[df['Channel Name'] == state]
    #         try:
    #             NewDf.to_excel(writer,sheet_name = state)
    #         except:
    #             NewDf.to_excel(writer,sheet_name = state[:31])
            
    # writer.save()
    #df.drop(columns=['Channel Name'], inplace=True)
    # Channel_Product_WiseReport

    # Product Wise Group Report
    final_total=product_grp['Status'].count()
    semi_total = product_grp['Status'].count() - product_grp['Status'].apply(lambda x:x.str.contains('CANCEL').sum())

    status_delivered = product_grp['Status'].apply(lambda x:x.str.contains('DELIVERED').sum())
    status_rto = product_grp['Status'].apply(lambda x:x.str.contains('RTO').sum())
    status_dto = product_grp['Status'].apply(lambda x:x.str.contains('DTO').sum())
    status_shipped = product_grp['Status'].apply(lambda x:x.str.contains('SHIPPED').sum())
    status_cancel = product_grp['Status'].apply(lambda x:x.str.contains('CANCEL').sum())

    total_profit = product_grp['Profit'].sum()
    tf=df['Profit'].sum()

    comboStatus_df = pd.concat([semi_total,status_delivered,status_rto,status_dto,status_shipped,status_cancel,final_total, total_profit],axis='columns',sort=False)
    comboStatus_df.set_axis(["SemiTotal", "DeliveredTotal", 'RtoTotal', 'DtoTotal', 'ShippedTotal', 'CancelTotal', 'TotalOrders','TotalProfit'], axis="columns",inplace=True)

    # Percentage Calculation
    comboStatus_df['PctProfit']=(comboStatus_df['TotalProfit']/tf)*100
    comboStatus_df['PctDelivered']=(comboStatus_df['DeliveredTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctRto']=(comboStatus_df['RtoTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctDto']=(comboStatus_df['DtoTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctShipped']=(comboStatus_df['ShippedTotal']/comboStatus_df['SemiTotal'])*100

    comboStatus_df['PctProfit'] = comboStatus_df['PctProfit'].round(2)
    comboStatus_df['PctDelivered'] = comboStatus_df['PctDelivered'].round(2)
    comboStatus_df['PctRto'] = comboStatus_df['PctRto'].round(2)
    comboStatus_df['PctDto'] = comboStatus_df['PctDto'].round(2)
    comboStatus_df['PctShipped'] = comboStatus_df['PctShipped'].round(2)
    
    comboStatus_df['PctProfit'] = comboStatus_df['PctProfit'].astype(str) + " %"
    comboStatus_df['PctDelivered'] = comboStatus_df['PctDelivered'].astype(str) + " %"
    comboStatus_df['PctRto'] = comboStatus_df['PctRto'].astype(str) + " %"
    comboStatus_df['PctDto'] = comboStatus_df['PctDto'].astype(str) + " %"
    comboStatus_df['PctShipped'] = comboStatus_df['PctShipped'].astype(str) + " %"


    Product_WiseReport = comboStatus_df.loc[:,['TotalProfit','PctProfit','TotalOrders','DeliveredTotal','PctDelivered', 'RtoTotal', 'PctRto','DtoTotal','PctDto','ShippedTotal','PctShipped','CancelTotal' ]]
    Product_WiseReport =  Product_WiseReport.sort_values(by=['TotalProfit'])
    
    # Product_WiseReport

    # Shipping Company Wise Group Report
    final_total=shipping_company_grp['Status'].count()
    semi_total = shipping_company_grp['Status'].count() - shipping_company_grp['Status'].apply(lambda x:x.str.contains('CANCEL').sum())

    status_delivered = shipping_company_grp['Status'].apply(lambda x:x.str.contains('DELIVERED').sum())
    status_rto = shipping_company_grp['Status'].apply(lambda x:x.str.contains('RTO').sum())
    status_dto = shipping_company_grp['Status'].apply(lambda x:x.str.contains('DTO').sum())
    status_shipped = shipping_company_grp['Status'].apply(lambda x:x.str.contains('SHIPPED').sum())
    status_cancel = shipping_company_grp['Status'].apply(lambda x:x.str.contains('CANCEL').sum())

    total_profit = shipping_company_grp['Profit'].sum()
    tf=df['Profit'].sum()

    comboStatus_df = pd.concat([semi_total,status_delivered,status_rto,status_dto,status_shipped,status_cancel,final_total, total_profit],axis='columns',sort=False)
    comboStatus_df.set_axis(["SemiTotal", "DeliveredTotal", 'RtoTotal', 'DtoTotal', 'ShippedTotal', 'CancelTotal', 'TotalOrders','TotalProfit'], axis="columns",inplace=True)

    # Percentage Calculation
    comboStatus_df['PctProfit']=(comboStatus_df['TotalProfit']/tf)*100
    comboStatus_df['PctDelivered']=(comboStatus_df['DeliveredTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctRto']=(comboStatus_df['RtoTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctDto']=(comboStatus_df['DtoTotal']/comboStatus_df['SemiTotal'])*100
    comboStatus_df['PctShipped']=(comboStatus_df['ShippedTotal']/comboStatus_df['SemiTotal'])*100

    comboStatus_df['PctProfit'] = comboStatus_df['PctProfit'].round(2)
    comboStatus_df['PctDelivered'] = comboStatus_df['PctDelivered'].round(2)
    comboStatus_df['PctRto'] = comboStatus_df['PctRto'].round(2)
    comboStatus_df['PctDto'] = comboStatus_df['PctDto'].round(2)
    comboStatus_df['PctShipped'] = comboStatus_df['PctShipped'].round(2)
    
    comboStatus_df['PctProfit'] = comboStatus_df['PctProfit'].astype(str) + " %"
    comboStatus_df['PctDelivered'] = comboStatus_df['PctDelivered'].astype(str) + " %"
    comboStatus_df['PctRto'] = comboStatus_df['PctRto'].astype(str) + " %"
    comboStatus_df['PctDto'] = comboStatus_df['PctDto'].astype(str) + " %"
    comboStatus_df['PctShipped'] = comboStatus_df['PctShipped'].astype(str) + " %"


    Shipping_Company_WiseReport = comboStatus_df.loc[:,['TotalProfit','PctProfit','TotalOrders','DeliveredTotal','PctDelivered', 'RtoTotal', 'PctRto','DtoTotal','PctDto','ShippedTotal','PctShipped','CancelTotal' ]]
    Shipping_Company_WiseReport = Shipping_Company_WiseReport.sort_values(by=['TotalProfit'])
    # Shipping_Company_WiseReport

    # Channel & Date Wise Report
    df2 = df.set_index('Order Date')
    channel_grp2 = df2.groupby('Channel Name')
    profitSum=channel_grp2['Profit'].resample('D').sum()
    orderCount=channel_grp2['Profit'].resample('D').count()
    channel_date_wise_df = pd.concat([profitSum,orderCount],axis='columns',sort=False)
    channel_date_wise_df.set_axis(['TotalProfit','TotalOrder'], axis="columns",inplace=True)
    # channel_date_wise_df

    # Save all Reports in Excel File
    print("Cols are", Channel_Product_WiseReport.columns)
    with pd.ExcelWriter('Output101.xlsx') as writer:
        ChannelWiseReport.to_excel(writer, sheet_name='Channel_Wise')
        ChannelWiseReport2.to_excel(writer, sheet_name='Channel_Wise2')
        Channel_Product_WiseReport.to_excel(writer, sheet_name='Channel_Product_Wise')
        Product_WiseReport.to_excel(writer, sheet_name='Product_Wise')
        Shipping_Company_WiseReport.to_excel(writer, sheet_name='Shipping_Company_Wise')
        channel_date_wise_df.to_excel(writer, sheet_name='channel_date_wise')

        # Plot Profit by Channel Chart
        profit_by__channel_chart = channel_grp['Profit'].sum()
        profit_by__channel_chart.plot.bar(x="Channel Name", y="TotalProfit", rot=80,
                                        title="Profit by Channel - Month Nov 2022") 

# run("C:/Users/joy/Desktop/smallData.xlsx")
