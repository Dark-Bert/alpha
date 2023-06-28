import numpy as np
import xlrd


def insertion(sheet, sheet_list, to_display):
    if sheet != mean:
        for i in range(sheet.nrows):
            for j in range(sheet.ncols):
                if i != 0:
                    sheet_list[i][j] = float(sheet.cell_value(i, j))

    else:
        for i in range(sheet.nrows):
            for j in range(sheet.ncols):
                if (j != 0) & (i != 0):
                    sheet_list[i][j] = float(sheet.cell_value(i, j))

    if to_display:
        for i in range(sheet.nrows):
            for j in range(sheet.ncols):
                print(str(sheet.cell_value(i, j)), end='\t')
            print('\n')
    return True  


def car_brands():
    for i in range(mean.nrows):
        print(mean.cell_value(i, 0))
    return True


def library(car_name, cc_weight_):
    match car_name:
        case 'sheet1':
            return beta_function(mean, mean_list, cc_weight_)
        case 'general':
            return beta_function(general, general_list, cc_weight_)
        case 'hyundai':
            return beta_function(hyundai, hyundai_list, cc_weight_)
        case 'kia':
            return beta_function(kia, kia_list, cc_weight_)
        case 'hyundai-kia':
            return beta_function(hyundai_kia, hyundai_kia_list, cc_weight_)
        case 'hyundai/kia':
            return beta_function(hyundai_kia, hyundai_kia_list, cc_weight_)
        case 'jeep':
            return beta_function(jeep, jeep_list, cc_weight_)
        case 'ford':
            return beta_function(ford, ford_list, cc_weight_)
        case 'suzuki':
            return beta_function(suzuki, suzuki_list, cc_weight_)
        case 'toyota':
            return beta_function(toyota, toyota_list, cc_weight_)
        case 'mazda':
            return beta_function(mazda, mazda_list, cc_weight_)
        case 'nissan':
            return beta_function(nissan, nissan_list, cc_weight_)
        case 'mitsubishi':
            return beta_function(mitsubishi, mitsubishi_list, cc_weight_)
        case 'honda':
            return beta_function(honda, honda_list, cc_weight_)
        case 'skoda':
            return beta_function(skoda, skoda_list, cc_weight_)
        case 'mercedes':
            return beta_function(mercedes, mercedes_list, cc_weight_)
        case 'mercedes-benz':
            return beta_function(mercedes, mercedes_list, cc_weight_)
        case 'renault':
            return beta_function(renault, renault_list, cc_weight_)
        case 'volkswagen':
            return beta_function(volkswagen, volkswagen_list, cc_weight_)
        case 'volkswagen-audi':
            return beta_function(volkswagen_audi, volkswagen_audi_list, cc_weight_)
        case 'volkswagen/audi':
            return beta_function(volkswagen_audi, volkswagen_audi_list, cc_weight_)
        case 'audi':
            return beta_function(audi, audi_list, cc_weight_)
        case 'bmw':
            return beta_function(bmw, bmw_list, cc_weight_)
        case 'fiat':
            return beta_function(fiat, fiat_list, cc_weight_)
        case 'volvo':
            return beta_function(volvo, volvo_list, cc_weight_)
        case 'opel':
            return beta_function(opel, opel_list, cc_weight_)
        case 'psa':
            return beta_function(psa, psa_list, cc_weight_)
        case _:
            car_name = str(input("Car brand not registered in record!, retry: "))
            return library(car_name, cc_weight_)


def beta_function(sheet, sheet_list, cc_weight_):  # the last row is a mean value row
    if not(np.array_equal(sheet_list, mean_list)):
        if cc_weight_ > sheet_list[sheet.nrows - 2][0]:
            cc_weight_ = float(input("Entered weight exceeded maximum value(" +
                                     str(sheet_list[sheet.nrows-2][0]) + ")! retry: "))
            return beta_function(sheet, sheet_list, cc_weight_)

        else:
            for k in range(sheet.nrows - 1):
                if (k >= 2) and (k <= (sheet.nrows - 2)):
                    if (cc_weight_ <= sheet_list[k][0]) and (cc_weight_ >= sheet_list[k - 1][0]):
                        f0 = abs(sheet_list[k][0] - cc_weight_)
                        f1 = abs(sheet_list[k - 1][0] - cc_weight_)
                        estimate_Pt = (f0*sheet_list[k][3] + f1*sheet_list[k - 1][3]) / (f0 + f1)
                        estimate_Pd = (f0*sheet_list[k][4] + f1*sheet_list[k - 1][4]) / (f0 + f1)
                        estimate_Rh = (f0*sheet_list[k][5] + f1*sheet_list[k - 1][5]) / (f0 + f1)
                        percent_Pt = (estimate_Pt / cc_weight_) * 100
                        percent_Pd = (estimate_Pd / cc_weight_) * 100
                        percent_Rh = (estimate_Rh / cc_weight_) * 100
                        print("Estimated (Pt, Pd, Rh) in percent: " + "(" +
                              str(round(percent_Pt, 3)) + ", " + str(round(percent_Pd, 3)) + ", "
                              + str(round(percent_Rh, 3)) + ")")
                        break

                    elif k == sheet.nrows - 2:
                        cc_weight_ = float(input("CC ceramic weight must be >= " + str(sheet_list[1][0] + ": ")))
                        return beta_function(sheet, sheet_list, cc_weight_)

        name_ = str(input("Enter the car brand name (lower case): "))
        if name_ != 'sheet1':
            cc_weight_ = float(input("Enter the weight of the ceramic block's weight (in grams): "))
        else:
            cc_weight_ = None
        return library(name_, cc_weight_)

    else:
        name_ = str(input("Enter the car brand name from 'sheet1' (lower case): "))
        match name_:
            case 'general':
                print("For expected weight (in grams): " + str(sheet_list[1][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[1][4]) + ", " +
                      str(sheet_list[1][5]) + ", " + str(sheet_list[1][6]))
                print("PGMs mass (in grams): " + str(sheet_list[1][7]) +
                      " with percent of: " + str(sheet_list[1][8]))
                print("Percent Pt: " + str((sheet_list[1][4]/sheet_list[1][7])*100))
                print("Percent Pd: " + str((sheet_list[1][5]/sheet_list[1][7])*100))
                print("Percent Rh: " + str((sheet_list[1][6]/sheet_list[1][7])*100))
                print('\n')
            case 'hyundai':
                print("For expected weight (in grams): " + str(sheet_list[2][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[2][4]) + ", " +
                      str(sheet_list[2][5]) + ", " + str(sheet_list[2][6]))
                print("PGMs mass (in grams): " + str(sheet_list[2][7]) +
                      " with percent of: " + str(sheet_list[2][8]))
                print("Percent Pt: " + str((sheet_list[2][4]/sheet_list[2][7])*100))
                print("Percent Pd: " + str((sheet_list[2][5]/sheet_list[2][7])*100))
                print("Percent Rh: " + str((sheet_list[2][6]/sheet_list[2][7])*100))
                print('\n')

            case 'kia':
                print("For expected weight (in grams): " + str(sheet_list[3][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[3][4]) + ", " +
                      str(sheet_list[3][5]) + ", " + str(sheet_list[3][6]))
                print("PGMs mass (in grams): " + str(sheet_list[3][7]) +
                      " with percent of: " + str(sheet_list[3][8]))
                print("Percent Pt: " + str((sheet_list[3][4]/sheet_list[3][7])*100))
                print("Percent Pd: " + str((sheet_list[3][5]/sheet_list[3][7])*100))
                print("Percent Rh: " + str((sheet_list[3][6]/sheet_list[3][7])*100))
                print('\n')

            case 'hyundai-kia':
                print("For expected weight (in grams): " + str(sheet_list[4][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[4][4]) + ", " +
                      str(sheet_list[4][5]) + ", " + str(sheet_list[4][6]))
                print("PGMs mass (in grams): " + str(sheet_list[4][7]) +
                      " with percent of: " + str(sheet_list[4][8]))
                print("Percent Pt: " + str((sheet_list[4][4]/sheet_list[4][7])*100))
                print("Percent Pd: " + str((sheet_list[4][5]/sheet_list[4][7])*100))
                print("Percent Rh: " + str((sheet_list[4][6]/sheet_list[4][7])*100))
                print('\n')

            case 'hyundai/kia':
                print("For expected weight (in grams): " + str(sheet_list[4][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[4][4]) + ", " +
                      str(sheet_list[4][5]) + ", " + str(sheet_list[4][6]))
                print("PGMs mass (in grams): " + str(sheet_list[4][7]) +
                      " with percent of: " + str(sheet_list[4][8]))
                print("Percent Pt: " + str((sheet_list[4][4]/sheet_list[4][7])*100))
                print("Percent Pd: " + str((sheet_list[4][5]/sheet_list[4][7])*100))
                print("Percent Rh: " + str((sheet_list[4][6]/sheet_list[4][7])*100))
                print('\n')

            case 'jeep':
                print("For expected weight (in grams): " + str(sheet_list[5][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[5][4]) + ", " +
                      str(sheet_list[5][5]) + ", " + str(sheet_list[5][6]))
                print("PGMs mass (in grams): " + str(sheet_list[5][7]) +
                      " with percent of: " + str(sheet_list[5][8]))
                print("Percent Pt: " + str((sheet_list[5][4]/sheet_list[5][7])*100))
                print("Percent Pd: " + str((sheet_list[5][5]/sheet_list[5][7])*100))
                print("Percent Rh: " + str((sheet_list[5][6]/sheet_list[5][7])*100))
                print('\n')

            case 'ford':
                print("For expected weight (in grams): " + str(sheet_list[6][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[6][4]) + ", " +
                      str(sheet_list[6][5]) + ", " + str(sheet_list[6][6]))
                print("PGMs mass (in grams): " + str(sheet_list[6][7]) +
                      " with percent of: " + str(sheet_list[6][8]))
                print("Percent Pt: " + str((sheet_list[6][4]/sheet_list[6][7])*100))
                print("Percent Pd: " + str((sheet_list[6][5]/sheet_list[6][7])*100))
                print("Percent Rh: " + str((sheet_list[6][6]/sheet_list[6][7])*100))
                print('\n')

            case 'suzuki':
                print("For expected weight (in grams): " + str(sheet_list[7][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[7][4]) + ", " +
                      str(sheet_list[7][5]) + ", " + str(sheet_list[7][6]))
                print("PGMs mass (in grams): " + str(sheet_list[7][7]) +
                      " with percent of: " + str(sheet_list[7][8]))
                print("Percent Pt: " + str((sheet_list[7][4]/sheet_list[7][7])*100))
                print("Percent Pd: " + str((sheet_list[7][5]/sheet_list[7][7])*100))
                print("Percent Rh: " + str((sheet_list[7][6]/sheet_list[7][7])*100))
                print('\n')

            case 'toyota':
                print("For expected weight (in grams): " + str(sheet_list[8][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[8][4]) + ", " +
                      str(sheet_list[8][5]) + ", " + str(sheet_list[8][6]))
                print("PGMs mass (in grams): " + str(sheet_list[8][7]) +
                      " with percent of: " + str(sheet_list[8][8]))
                print("Percent Pt: " + str((sheet_list[8][4]/sheet_list[8][7])*100))
                print("Percent Pd: " + str((sheet_list[8][5]/sheet_list[8][7])*100))
                print("Percent Rh: " + str((sheet_list[8][6]/sheet_list[8][7])*100))
                print('\n')

            case 'mazda':
                print("For expected weight (in grams): " + str(sheet_list[9][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[9][4]) + ", " +
                      str(sheet_list[9][5]) + ", " + str(sheet_list[9][6]))
                print("PGMs mass (in grams): " + str(sheet_list[9][7]) +
                      " with percent of: " + str(sheet_list[9][8]))
                print("Percent Pt: " + str((sheet_list[9][4]/sheet_list[9][7])*100))
                print("Percent Pd: " + str((sheet_list[9][5]/sheet_list[9][7])*100))
                print("Percent Rh: " + str((sheet_list[9][6]/sheet_list[9][7])*100))
                print('\n')

            case 'nissan':
                print("For expected weight (in grams): " + str(sheet_list[10][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[10][4]) + ", " +
                      str(sheet_list[10][5]) + ", " + str(sheet_list[10][6]))
                print("PGMs mass (in grams): " + str(sheet_list[11][7]) +
                      " with percent of: " + str(sheet_list[10][8]))
                print("Percent Pt: " + str((sheet_list[10][4]/sheet_list[10][7])*100))
                print("Percent Pd: " + str((sheet_list[10][5]/sheet_list[10][7])*100))
                print("Percent Rh: " + str((sheet_list[10][6]/sheet_list[10][7])*100))
                print('\n')

            case 'mitsubishi':
                print("For expected weight (in grams): " + str(sheet_list[11][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[11][4]) + ", " +
                      str(sheet_list[11][5]) + ", " + str(sheet_list[11][6]))
                print("PGMs mass (in grams): " + str(sheet_list[11][7]) +
                      " with percent of: " + str(sheet_list[11][8]))
                print("Percent Pt: " + str((sheet_list[11][4]/sheet_list[11][7])*100))
                print("Percent Pd: " + str((sheet_list[11][5]/sheet_list[11][7])*100))
                print("Percent Rh: " + str((sheet_list[11][6]/sheet_list[11][7])*100))
                print('\n')

            case 'honda':
                print("For expected weight (in grams): " + str(sheet_list[12][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[12][4]) + ", " +
                      str(sheet_list[12][5]) + ", " + str(sheet_list[12][6]))
                print("PGMs mass (in grams): " + str(sheet_list[12][7]) +
                      " with percent of: " + str(sheet_list[12][8]))
                print("Percent Pt: " + str((sheet_list[12][4]/sheet_list[12][7])*100))
                print("Percent Pd: " + str((sheet_list[12][5]/sheet_list[12][7])*100))
                print("Percent Rh: " + str((sheet_list[12][6]/sheet_list[12][7])*100))
                print('\n')

            case 'skoda':
                print("For expected weight (in grams): " + str(sheet_list[13][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[13][4]) + ", " +
                      str(sheet_list[13][5]) + ", " + str(sheet_list[13][6]))
                print("PGMs mass (in grams): " + str(sheet_list[13][7]) +
                      " with percent of: " + str(sheet_list[13][8]))
                print("Percent Pt: " + str((sheet_list[13][4]/sheet_list[13][7])*100))
                print("Percent Pd: " + str((sheet_list[13][5]/sheet_list[13][7])*100))
                print("Percent Rh: " + str((sheet_list[13][6]/sheet_list[13][7])*100))
                print('\n')

            case 'mercedes':
                print("For expected weight (in grams): " + str(sheet_list[14][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[14][4]) + ", " +
                      str(sheet_list[14][5]) + ", " + str(sheet_list[14][6]))
                print("PGMs mass (in grams): " + str(sheet_list[14][7]) +
                      " with percent of: " + str(sheet_list[14][8]))
                print("Percent Pt: " + str((sheet_list[14][4]/sheet_list[14][7])*100))
                print("Percent Pd: " + str((sheet_list[14][5]/sheet_list[14][7])*100))
                print("Percent Rh: " + str((sheet_list[14][6]/sheet_list[14][7])*100))
                print('\n')

            case 'mercedes-benz':
                print("For expected weight (in grams): " + str(sheet_list[14][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[14][4]) + ", " +
                      str(sheet_list[14][5]) + ", " + str(sheet_list[14][6]))
                print("PGMs mass (in grams): " + str(sheet_list[14][7]) +
                      " with percent of: " + str(sheet_list[14][8]))
                print("Percent Pt: " + str((sheet_list[14][4]/sheet_list[14][7])*100))
                print("Percent Pd: " + str((sheet_list[14][5]/sheet_list[14][7])*100))
                print("Percent Rh: " + str((sheet_list[14][6]/sheet_list[14][7])*100))
                print('\n')

            case 'renault':
                print("For expected weight (in grams): " + str(sheet_list[15][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[15][4]) + ", " +
                      str(sheet_list[15][5]) + ", " + str(sheet_list[15][6]))
                print("PGMs mass (in grams): " + str(sheet_list[15][7]) +
                      " with percent of: " + str(sheet_list[15][8]))
                print("Percent Pt: " + str((sheet_list[15][4]/sheet_list[15][7])*100))
                print("Percent Pd: " + str((sheet_list[15][5]/sheet_list[15][7])*100))
                print("Percent Rh: " + str((sheet_list[15][6]/sheet_list[15][7])*100))
                print('\n')

            case 'volkswagen':
                print("For expected weight (in grams): " + str(sheet_list[16][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[16][4]) + ", " +
                      str(sheet_list[16][5]) + ", " + str(sheet_list[16][6]))
                print("PGMs mass (in grams): " + str(sheet_list[16][7]) +
                      " with percent of: " + str(sheet_list[16][8]))
                print("Percent Pt: " + str((sheet_list[16][4]/sheet_list[16][7])*100))
                print("Percent Pd: " + str((sheet_list[16][5]/sheet_list[16][7])*100))
                print("Percent Rh: " + str((sheet_list[16][6]/sheet_list[16][7])*100))
                print('\n')

            case 'volkswagen-audi':
                print("For expected weight (in grams): " + str(sheet_list[17][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[17][4]) + ", " +
                      str(sheet_list[17][5]) + ", " + str(sheet_list[17][6]))
                print("PGMs mass (in grams): " + str(sheet_list[17][7]) +
                      " with percent of: " + str(sheet_list[17][8]))
                print("Percent Pt: " + str((sheet_list[17][4]/sheet_list[17][7])*100))
                print("Percent Pd: " + str((sheet_list[17][5]/sheet_list[17][7])*100))
                print("Percent Rh: " + str((sheet_list[17][6]/sheet_list[17][7])*100))
                print('\n')

            case 'volkswagen/audi':
                print("For expected weight (in grams): " + str(sheet_list[17][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[17][4]) + ", " +
                      str(sheet_list[17][5]) + ", " + str(sheet_list[17][6]))
                print("PGMs mass (in grams): " + str(sheet_list[17][7]) +
                      " with percent of: " + str(sheet_list[17][8]))
                print("Percent Pt: " + str((sheet_list[17][4]/sheet_list[17][7])*100))
                print("Percent Pd: " + str((sheet_list[17][5]/sheet_list[17][7])*100))
                print("Percent Rh: " + str((sheet_list[17][6]/sheet_list[17][7])*100))
                print('\n')

            case 'audi':
                print("For expected weight (in grams): " + str(sheet_list[19][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[18][4]) + ", " +
                      str(sheet_list[18][5]) + ", " + str(sheet_list[18][6]))
                print("PGMs mass (in grams): " + str(sheet_list[21][7]) +
                      " with percent of: " + str(sheet_list[18][8]))
                print("Percent Pt: " + str((sheet_list[18][4]/sheet_list[18][7])*100))
                print("Percent Pd: " + str((sheet_list[18][5]/sheet_list[18][7])*100))
                print("Percent Rh: " + str((sheet_list[18][6]/sheet_list[18][7])*100))
                print('\n')

            case 'bmw':
                print("For expected weight (in grams): " + str(sheet_list[19][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[19][4]) + ", " +
                      str(sheet_list[19][5]) + ", " + str(sheet_list[19][6]))
                print("PGMs mass (in grams): " + str(sheet_list[19][7]) +
                      " with percent of: " + str(sheet_list[19][8]))
                print("Percent Pt: " + str((sheet_list[19][4]/sheet_list[19][7])*100))
                print("Percent Pd: " + str((sheet_list[19][5]/sheet_list[19][7])*100))
                print("Percent Rh: " + str((sheet_list[19][6]/sheet_list[19][7])*100))
                print('\n')

            case 'fiat':
                print("For expected weight (in grams): " + str(sheet_list[20][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[20][4]) + ", " +
                      str(sheet_list[20][5]) + ", " + str(sheet_list[20][6]))
                print("PGMs mass (in grams): " + str(sheet_list[20][7]) +
                      " with percent of: " + str(sheet_list[20][8]))
                print("Percent Pt: " + str((sheet_list[20][4]/sheet_list[20][7])*100))
                print("Percent Pd: " + str((sheet_list[20][5]/sheet_list[20][7])*100))
                print("Percent Rh: " + str((sheet_list[20][6]/sheet_list[20][7])*100))
                print('\n')

            case 'volvo':
                print("For expected weight (in grams): " + str(sheet_list[21][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[21][4]) + ", " +
                      str(sheet_list[21][5]) + ", " + str(sheet_list[21][6]))
                print("PGMs mass (in grams): " + str(sheet_list[21][7]) +
                      " with percent of: " + str(sheet_list[21][8]))
                print("Percent Pt: " + str((sheet_list[21][4]/sheet_list[21][7])*100))
                print("Percent Pd: " + str((sheet_list[21][5]/sheet_list[21][7])*100))
                print("Percent Rh: " + str((sheet_list[21][6]/sheet_list[21][7])*100))
                print('\n')

            case 'opel':
                print("For expected weight (in grams): " + str(sheet_list[22][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[22][4]) + ", " +
                      str(sheet_list[22][5]) + ", " + str(sheet_list[22][6]))
                print("PGMs mass (in grams): " + str(sheet_list[22][7]) +
                      " with percent of: " + str(sheet_list[22][8]))
                print("Percent Pt: " + str((sheet_list[22][4]/sheet_list[22][7])*100))
                print("Percent Pd: " + str((sheet_list[22][5]/sheet_list[22][7])*100))
                print("Percent Rh: " + str((sheet_list[22][6]/sheet_list[22][7])*100))
                print('\n')

            case 'psa':
                print("For expected weight (in grams): " + str(sheet_list[23][1]))
                print("(Pt, Pd, Rh) (in grams): " + str(sheet_list[23][4]) + ", " +
                      str(sheet_list[23][5]) + ", " + str(sheet_list[23][6]))
                print("PGMs mass (in grams): " + str(sheet_list[23][7]) +
                      " with percent of: " + str(sheet_list[23][8]))
                print("Percent Pt: " + str((sheet_list[23][4]/sheet_list[23][7])*100))
                print("Percent Pd: " + str((sheet_list[23][5]/sheet_list[23][7])*100))
                print("Percent Rh: " + str((sheet_list[23][6]/sheet_list[23][7])*100))
                print('\n')

            case _:
                name_ = str(input("Car brand not registered in record!, type 'sheet1' to continue: "))
                if name_ == 'sheet1':
                    return beta_function(mean, mean_list, None)

        name_ = str(input("Do you want to exit from 'sheet1'? (y/n): "))

        if name_ == 'y' or name_ == 'Y':
            name_ = str(input("Do you want to exit from the program? (press '!' key): "))
            if name_ == '!':
                print("Exited from the program successfully!")
                exit(0)
            name_ = str(input("Enter the car brand name (lower case): "))
            cc_weight_ = float(input("Enter the weight of the ceramic block's weight (in grams): "))
            return library(name_, cc_weight_)

        elif name_ == 'n' or name_ == 'N':
            return beta_function(mean, mean_list, None)


directory = 'Data analysis of CC mesh PGMs percent in different automotive brands.xlsx'
this_file = xlrd.open_workbook(directory)

mean = this_file.sheet_by_name('Sheet1')
mean_list = np.zeros((mean.nrows, mean.ncols))

general = this_file.sheet_by_name('General')
general_list = np.zeros((general.nrows, general.ncols))

hyundai = this_file.sheet_by_name('Hyundai')
hyundai_list = np.zeros((hyundai.nrows, hyundai.ncols))

kia = this_file.sheet_by_name('KIA')
kia_list = np.zeros((kia.nrows, kia.ncols))

hyundai_kia = this_file.sheet_by_name('Hyundai_KIA')
hyundai_kia_list = np.zeros((hyundai_kia.nrows, hyundai_kia.ncols))

jeep = this_file.sheet_by_name('Jeep')
jeep_list = np.zeros((jeep.nrows, jeep.ncols))

ford = this_file.sheet_by_name('Ford')
ford_list = np.zeros((ford.nrows, ford.ncols))

suzuki = this_file.sheet_by_name('Suzuki')
suzuki_list = np.zeros((suzuki.nrows, suzuki.ncols))

toyota = this_file.sheet_by_name('Toyota')
toyota_list = np.zeros((toyota.nrows, toyota.ncols))

mazda = this_file.sheet_by_name('Mazda')
mazda_list = np.zeros((mazda.nrows, mazda.ncols))

nissan = this_file.sheet_by_name('Nissan')
nissan_list = np.zeros((nissan.nrows, nissan.ncols))

mitsubishi = this_file.sheet_by_name('Mitsubishi')
mitsubishi_list = np.zeros((mitsubishi.nrows, mitsubishi.ncols))

honda = this_file.sheet_by_name('Honda')
honda_list = np.zeros((honda.nrows, honda.ncols))

skoda = this_file.sheet_by_name('Skoda')
skoda_list = np.zeros((skoda.nrows, skoda.ncols))

mercedes = this_file.sheet_by_name('Mercedes_Benz')
mercedes_list = np.zeros((mercedes.nrows, mercedes.ncols))

renault = this_file.sheet_by_name('Renault')
renault_list = np.zeros((renault.nrows, renault.ncols))

volkswagen = this_file.sheet_by_name('Volkswagen')
volkswagen_list = np.zeros((volkswagen.nrows, volkswagen.ncols))

volkswagen_audi = this_file.sheet_by_name('Volkswagen_Audi')
volkswagen_audi_list = np.zeros((volkswagen_audi.nrows, volkswagen_audi.ncols))

audi = this_file.sheet_by_name('Audi')
audi_list = np.zeros((audi.nrows, audi.ncols))

bmw = this_file.sheet_by_name('BMW')
bmw_list = np.zeros((bmw.nrows, bmw.ncols))

fiat = this_file.sheet_by_name('Fiat')
fiat_list = np.zeros((fiat.nrows, fiat.ncols))

volvo = this_file.sheet_by_name('volvo')
volvo_list = np.zeros((volvo.nrows, volvo.ncols))

opel = this_file.sheet_by_name('Opel')
opel_list = np.zeros((opel.nrows, opel.ncols))

psa = this_file.sheet_by_name('PSA')
psa_list = np.zeros((psa.nrows, psa.ncols))

insertion(mean, mean_list, 0)
insertion(general, general_list, 0)
insertion(hyundai, hyundai_list, 0)
insertion(kia, kia_list, 0)
insertion(hyundai_kia, hyundai_kia_list, 0)
insertion(jeep, jeep_list, 0)
insertion(ford, ford_list, 0)
insertion(suzuki, suzuki_list, 0)
insertion(toyota, toyota_list, 0)
insertion(mazda, mazda_list, 0)
insertion(nissan, nissan_list, 0)
insertion(mitsubishi, mitsubishi_list, 0)
insertion(honda, honda_list, 0)
insertion(skoda, skoda_list, 0)
insertion(mercedes, mercedes_list, 0)
insertion(renault, renault_list, 0)
insertion(volkswagen, volkswagen_list, 0)
insertion(volkswagen_audi, volkswagen_audi_list, 0)
insertion(audi, audi_list, 0)
insertion(bmw, bmw_list, 0)
insertion(fiat, fiat_list, 0)
insertion(volvo, volvo_list, 0)
insertion(opel, opel_list, 0)
insertion(psa, psa_list, 0)

car_brands()

print('\n'+'\t'+"To access mean data of all brands, just type 'sheet1'"+'\t'+'\n')

name = str(input("Enter the car brand name (lower case): "))

if name != 'sheet1':
    cc_weight = float(input("Enter the weight of the ceramic block's weight (in grams): "))
else:
    cc_weight = None

library(name, cc_weight)
