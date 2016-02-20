def prepare_sheet(worksheet):

    worksheet.set_landscape()
    worksheet.set_paper(9)
    #set_margins([left=0.7,] right=0.7,] top=0.75,] bottom=0.75]]])
    worksheet.set_margins(left=0.315, right=0.315, top=0.348, bottom=0.354)
    worksheet.center_horizontally()
    worksheet.set_print_scale(78)

    worksheet.set_column('A:BE', 2.4)
    worksheet.set_column('B:D', 2.1)
    worksheet.set_column('Q:Z', 2.1)
    worksheet.set_column('AM:AN', 2.6)
    worksheet.set_column('BA:BB', 2.6)

    worksheet.set_row(0, 14)
    #worksheet.set_row(start_row+3, 64)

    for n_row in range(12,175):
        worksheet.set_row(n_row, 12)
