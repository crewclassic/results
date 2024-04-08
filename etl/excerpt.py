event = ""
current_event = ""
for row in range(0,sheet.nrows):
    for col in range(0,16):
        # persist the event name
        if( row > 0 and col == 3):
            if(event == "" and sheet.cell_value(row,col) != ""):
                event = sheet.cell_value(row,col)
                current_event = event
            elif (event != "" and sheet.cell_value(row,col) == ""):
                current_event = event
            else:
                current_event = sheet.cell_value(row,col)
                event = current_event

    if(row == 0):
        print("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}".format(
            sheet.cell_value(row,0),
            sheet.cell_value(row,1),
            sheet.cell_value(row,2),
            sheet.cell_value(row,3),
            sheet.cell_value(row,4),
            sheet.cell_value(row,5),
            sheet.cell_value(row,6),
            sheet.cell_value(row,7),
            sheet.cell_value(row,8),
            sheet.cell_value(row,9),
            sheet.cell_value(row,10),
            sheet.cell_value(row,11),
            sheet.cell_value(row,12),
            sheet.cell_value(row,13),
            sheet.cell_value(row,14)))

    else:
        col0 = sheet.cell_value(row,0)
        col1 = sheet.cell_value(row,1)
        strtime = sheet.cell_value(row,5)
        try:
            col0 = int(col0)
        except:
            pass
        try:
            col1 = int(col1)
        except:
            pass
        
        try:
            x = int(sheet.cell_value(row,5) * 24 * 3600 * 100)  # milliseconds
            hour = x//360000
            min  = (x%360000)//6000
            sec  = x//100%60
            mlsec = x%100//1*10000
#            mlsec = round((x%100/10))*100000
#            mlsec = int(Decimal(x%100/10).to_integral_value(rounding=ROUND_HALF_UP))*100000
            my_time = time(hour, min, sec, mlsec)
            strtime = my_time.strftime('%M:%S.%f')[:-4]
        except:
            pass

        print("{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}".format(
            col0,
            col1,
            sheet.cell_value(row,2),
            current_event,
            sheet.cell_value(row,4),
            strtime,
            sheet.cell_value(row,6),
            sheet.cell_value(row,7),
            sheet.cell_value(row,8),
            sheet.cell_value(row,9),
            sheet.cell_value(row,10),
            sheet.cell_value(row,11),
            sheet.cell_value(row,12),
            sheet.cell_value(row,13),
            sheet.cell_value(row,14)))

        ---

            if(trophy == "" and row[1] != ""):
                trophy = row[1]
                current_trophy = trophy
            elif (trophy != "" and row[1] == ""):
                current_trophy = trophy
            else:
                current_trophy = row[1]
                trophy = current_trophy

            row[1] = current_trophy.strip()
