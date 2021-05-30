from gekko import GEKKO
from tkinter import *
import numpy as np
import pandas as pd
import os


def openfile():
    os.chdir('C:\\users\\evilm\\pycharmprojects\\pythonproject')
    os.system('start excel.exe График_работы_генераторов.xlsx')


def mat():
    price_for_start1 = type1_entry.get()
    price_for_start2 = type2_entry.get()
    price_for_start3 = type3_entry.get()
    price_for_hour1 = price1_entry.get()
    price_for_hour2 = price2_entry.get()
    price_for_hour3 = price3_entry.get()
    price_for_mw1 = pricemw1_entry.get()
    price_for_mw2 = pricemw2_entry.get()
    price_for_mw3 = pricemw3_entry.get()

    m = GEKKO(remote=False)

    x = [m.Var(value=1500, lb=1500, ub=4000) for i in range(25)]
    k = [m.Var(value=1, lb=0, ub=1, integer=True) for i in range(25)]

    y = [m.Var(value=1250, lb=1250, ub=1750) for i in range(50)]
    j = [m.Var(value=1, lb=0, ub=1, integer=True) for i in range(50)]

    z = [m.Var(value=850, lb=850, ub=2000) for i in range(60)]
    c = [m.Var(value=1, lb=0, ub=1, integer=True) for i in range(60)]

    m.Equation(
        x[0] * k[0] + x[1] * k[1] + x[2] * k[2] + x[3] * k[3] + x[4] * k[4] + y[0] * j[0] + y[1] * j[1]
        + y[2] * j[2] + y[3] * j[3] + y[4] * j[4] + y[5] * j[5] + y[6] * j[6] + y[7] * j[7] + y[8] * j[8]
        + y[9] * j[9] + z[0] * c[0] + z[1] * c[1] + z[2] * c[2] + z[3] * c[3] + z[4] * c[4] + z[5] * c[5]
        + z[6] * c[6] + z[7] * c[7] + z[8] * c[8] + z[9] * c[9] + z[10] * c[10] + z[11] * c[11] >= 17250)

    m.Equation(
        x[5] * k[5] + x[6] * k[6] + x[7] * k[7] + x[8] * k[8] + x[9] * k[9] + y[10] * j[10] + y[11] * j[11]
        + y[12] * j[12] + y[13] * j[13] + y[14] * j[14] + y[15] * j[15] + y[16] * j[16] + y[17] * j[17]
        + y[18] * j[18] + y[19] * j[19] + z[12] * c[12] + z[13] * c[13] + z[14] * c[14] + z[15] * c[15]
        + z[16] * c[16] + z[17] * c[17] + z[18] * c[18] + z[19] * c[19] + z[20] * c[20] + z[21] * c[21]
        + z[22] * c[22] + z[23] * c[23] >= 34500)

    m.Equation(
        x[10] * k[10] + x[11] * k[11] + x[12] * k[12] + x[13] * k[13] + x[14] * k[14] + y[20] * j[20]
        + y[21] * j[21] + y[22] * j[22] + y[23] * j[23] + y[24] * j[24] + y[25] * j[25] + y[26] * j[26]
        + y[27] * j[27] + y[28] * j[28] + y[29] * j[29] + z[24] * c[24] + z[25] * c[25] + z[26] * c[26]
        + z[27] * c[27] + z[28] * c[28] + z[29] * c[29] + z[30] * c[30] + z[31] * c[31] + z[32] * c[32]
        + z[33] * c[33] + z[34] * c[34] + z[35] * c[35] >= 28750)

    m.Equation(
        x[15] * k[15] + x[16] * k[16] + x[17] * k[17] + x[18] * k[18] + x[19] * k[19] + y[30] * j[30]
        + y[31] * j[31] + y[32] * j[32] + y[33] * j[33] + y[34] * j[34] + y[35] * j[35] + y[36] * j[36]
        + y[37] * j[37] + y[38] * j[38] + y[39] * j[39] + z[36] * c[36] + z[37] * c[37] + z[38] * c[38]
        + z[39] * c[39] + z[40] * c[40] + z[41] * c[41] + z[42] * c[42] + z[43] * c[43] + z[44] * c[44]
        + z[45] * c[45] + z[46] * c[46] + z[47] * c[47] >= 46000)

    m.Equation(
        x[20] * k[20] + x[21] * k[21] + x[22] * k[22] + x[23] * k[23] + x[24] * k[24] + y[40] * j[40]
        + y[41] * j[41] + y[42] * j[42] + y[43] * j[43] + y[44] * j[44] + y[45] * j[45] + y[46] * j[46]
        + y[47] * j[47] + y[48] * j[48] + y[49] * j[49] + z[48] * c[48] + z[49] * c[49] + z[50] * c[50]
        + z[51] * c[51] + z[52] * c[52] + z[53] * c[53] + z[54] * c[54] + z[55] * c[55] + z[56] * c[56]
        + z[57] * c[57] + z[58] * c[58] + z[59] * c[59] >= 31050)

    p1 = m.Param(int(price_for_start1))
    p2 = m.Param(int(price_for_start2))
    p3 = m.Param(int(price_for_start3))

    ph1 = m.Param(int(price_for_hour1))
    ph2 = m.Param(int(price_for_hour2))
    ph3 = m.Param(int(price_for_hour3))

    pm3 = m.Param(int(price_for_mw3))
    pm2 = m.Param(int(price_for_mw2))
    pm1 = m.Param(int(price_for_mw1))

    m.Obj(((x[0] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[0] + ((x[1] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[
        1] + (
                  (x[2] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[2] + (
                  (x[3] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[3] + (
                  (x[4] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[4] + (
                  (x[5] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[5] + (
                  (x[6] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[6] + (
                  (x[7] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[7] + (
                  (x[8] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[8] + (
                  (x[9] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[9] + (
                  (x[10] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[10] + (
                  (x[11] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[11] + (
                  (x[12] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[12] + (
                  (x[13] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[13] + (
                  (x[14] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[14] + (
                  (x[15] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[15] + (
                  (x[16] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[16] + (
                  (x[17] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[17] + (
                  (x[18] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[18] + (
                  (x[19] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[19] + (
                  (x[20] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[20] + (
                  (x[21] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[21] + (
                  (x[22] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[22] + (
                  (x[23] - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k[23] + (
                  (x[24] - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k[24] + (max(k[0].value - 0, 0) * p3.VALUE) + (
                  max(k[5].value - k[0].value, 0) * p3.VALUE) + (
                  max(k[10].value - k[5].value, 0) * p3.VALUE) + (max(k[15].value - k[10].value, 0) * p3.VALUE)
          + (max(k[20].value - k[15].value, 0) * p3.VALUE) + (max(k[1].value - 0, 0) * p3.VALUE) + (
                  max(k[6].value - k[1].value, 0) * p3.VALUE) + (
                  max(k[11].value - k[6].value, 0) * p3.VALUE) + (max(k[16].value - k[11].value, 0) * p3.VALUE) + (
                  max(k[21].value - k[16].value, 0) * p3.VALUE) + (max(k[2].value - 0, 0) * p3.VALUE) + (
                  max(k[7].value - k[2].value, 0) * p3.VALUE) + (
                  max(k[12].value - k[7].value, 0) * p3.VALUE) + (max(k[17].value - k[12].value, 0) * p3.VALUE) + (
                  max(k[22].value - k[17].value, 0) * p3.VALUE) + (max(k[3].value - 0, 0) * p3.VALUE) + (
                  max(k[8].value - k[3].value, 0) * p3.VALUE) + (
                  max(k[13].value - k[8].value, 0) * p3.VALUE) + (max(k[18].value - k[13].value, 0) * p3.VALUE) + (
                  max(k[23].value - k[18].value, 0) * p3.VALUE) + (max(k[4].value - 0, 0) * p3.VALUE) + (
                  max(k[9].value - k[4].value, 0) * p3.VALUE) + (
                  max(k[14].value - k[9].value, 0) * p3.VALUE) + (max(k[19].value - k[14].value, 0) * p3.VALUE) + (
                  max(k[24].value - k[19].value, 0) * p3.VALUE) + ((y[0] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[
              0] + (
                  (y[1] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[1] + (
                  (y[2] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[2] + (
                  (y[3] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[3] + (
                  (y[4] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[4] + (
                  (y[5] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[5] + (
                  (y[6] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[6] + (
                  (y[7] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[7] + (
                  (y[8] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[8] + (
                  (y[9] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[9] + (
                  (y[10] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[10] + (
                  (y[11] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[11] + (
                  (y[12] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[12] + (
                  (y[13] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[13] + (
                  (y[14] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[14] + (
                  (y[15] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[15] + (
                  (y[16] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[16] + (
                  (y[17] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[17] + (
                  (y[18] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[18] + (
                  (y[19] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[19] + (
                  (y[20] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[20] + (
                  (y[21] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[21] + (
                  (y[22] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[22] + (
                  (y[23] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[23] + (
                  (y[24] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[24] + (
                  (y[25] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[25] + (
                  (y[26] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[26] + (
                  (y[27] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[27] + (
                  (y[28] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[28] + (
                  (y[29] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[29] + (
                  (y[30] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[30] + (
                  (y[31] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[31] + (
                  (y[32] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[32] + (
                  (y[33] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[33] + (
                  (y[34] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[34] + (
                  (y[35] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[35] + (
                  (y[36] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[36] + (
                  (y[37] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[37] + (
                  (y[38] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[38] + (
                  (y[39] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[39] + (
                  (y[40] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[40] + (
                  (y[41] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[41] + (
                  (y[42] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 3) * j[42] + (
                  (y[43] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[43] + (
                  (y[44] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[44] + (
                  (y[45] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[45] + (
                  (y[46] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[46] + (
                  (y[47] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[47] + (
                  (y[48] - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j[48] + (
                  (y[49] - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j[49] + (max(j[0].value - 0, 0) * p2.VALUE) + (
                  max(j[10].value - j[0].value, 0) * p2.VALUE) + (
                  max(j[20].value - j[10].value, 0) * p2.VALUE) + (max(j[30].value - j[20].value, 0) * p2.VALUE)
          + (max(j[40].value - j[30].value, 0) * p2.VALUE) + (max(j[1].value - 0, 0) * p2.VALUE) + (
                  max(j[11].value - j[1].value, 0) * p2.VALUE) + (
                  max(j[21].value - j[11].value, 0) * p2.VALUE) + (max(j[31].value - j[21].value, 0) * p2.VALUE) + (
                  max(j[41].value - j[31].value, 0) * p2.VALUE) + (max(j[2].value - 0, 0) * p2.VALUE) + (
                  max(j[12].value - j[2].value, 0) * p2.VALUE) + (
                  max(j[22].value - j[12].value, 0) * p2.VALUE) + (max(j[32].value - j[22].value, 0) * p2.VALUE) + (
                  max(j[42].value - j[32].value, 0) * p2.VALUE) + (max(j[3].value - 0, 0) * p2.VALUE) + (
                  max(j[13].value - j[3].value, 0) * p2.VALUE) + (
                  max(j[23].value - j[13].value, 0) * p2.VALUE) + (max(j[33].value - j[23].value, 0) * p2.VALUE) + (
                  max(j[43].value - j[33].value, 0) * p2.VALUE) + (max(j[4].value - 0, 0) * p2.VALUE) + (
                  max(j[14].value - j[4].value, 0) * p2.VALUE) + (
                  max(j[24].value - j[14].value, 0) * p2.VALUE) + (max(j[34].value - j[24].value, 0) * p2.VALUE) + (
                  max(j[44].value - j[34].value, 0) * p2.VALUE) + (max(j[5].value - 0, 0) * p2.VALUE) + (
                  max(j[15].value - j[5].value, 0) * p2.VALUE) + (
                  max(j[25].value - j[15].value, 0) * p2.VALUE) + (max(j[35].value - j[25].value, 0) * p2.VALUE)
          + (max(j[45].value - j[35].value, 0) * p2.VALUE) + (max(j[6].value - 0, 0) * p2.VALUE) + (
                  max(j[16].value - j[6].value, 0) * p2.VALUE) + (
                  max(j[26].value - j[16].value, 0) * p2.VALUE) + (max(j[36].value - j[26].value, 0) * p2.VALUE) + (
                  max(j[46].value - j[36].value, 0) * p2.VALUE) + (max(j[7].value - 0, 0) * p2.VALUE) + (
                  max(j[17].value - j[7].value, 0) * p2.VALUE) + (
                  max(j[27].value - j[17].value, 0) * p2.VALUE) + (max(j[37].value - j[27].value, 0) * p2.VALUE) + (
                  max(j[47].value - j[37].value, 0) * p2.VALUE) + (max(j[8].value - 0, 0) * p2.VALUE) + (
                  max(j[18].value - j[8].value, 0) * p2.VALUE) + (
                  max(j[28].value - j[18].value, 0) * p2.VALUE) + (max(j[38].value - j[28].value, 0) * p2.VALUE) + (
                  max(j[48].value - j[38].value, 0) * p2.VALUE) + (max(j[9].value - 0, 0) * p2.VALUE) + (
                  max(j[19].value - j[9].value, 0) * p2.VALUE) + (
                  max(j[29].value - j[19].value, 0) * p2.VALUE) + (max(j[39].value - j[29].value, 0) * p2.VALUE) + (
                  max(j[49].value - j[39].value, 0) * p2.VALUE) + ((z[0] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[
              0] + (
                  (z[1] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[1] + (
                  (z[2] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[2] + (
                  (z[3] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[3] + (
                  (z[4] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[4] + (
                  (z[5] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[5] + (
                  (z[6] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[6] + (
                  (z[7] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[7] + (
                  (z[8] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[8] + (
                  (z[9] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[9] + (
                  (z[10] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[10] + (
                  (z[11] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[11] + (
                  (z[12] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[12] + (
                  (z[13] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[13] + (
                  (z[14] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[14] + (
                  (z[15] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[15] + (
                  (z[16] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[16] + (
                  (z[17] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[17] + (
                  (z[18] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[18] + (
                  (z[19] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[19] + (
                  (z[20] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[20] + (
                  (z[21] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[21] + (
                  (z[22] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[22] + (
                  (z[23] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[23] + (
                  (z[24] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[24] + (
                  (z[25] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[25] + (
                  (z[26] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[26] + (
                  (z[27] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[27] + (
                  (z[28] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[28] + (
                  (z[29] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[29] + (
                  (z[30] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[30] + (
                  (z[31] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[31] + (
                  (z[32] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[32] + (
                  (z[33] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[33] + (
                  (z[34] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[34] + (
                  (z[35] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[35] + (
                  (z[36] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[36] + (
                  (z[37] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[37] + (
                  (z[38] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[38] + (
                  (z[39] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[39] + (
                  (z[40] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[40] + (
                  (z[41] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[41] + (
                  (z[42] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[42] + (
                  (z[43] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[43] + (
                  (z[44] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[44] + (
                  (z[45] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[45] + (
                  (z[46] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[46] + (
                  (z[47] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[47] + (
                  (z[48] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[48] + (
                  (z[49] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[49] + (
                  (z[50] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[50] + (
                  (z[51] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[51] + (
                  (z[52] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[52] + (
                  (z[53] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[53] + (
                  (z[54] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[54] + (
                  (z[55] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[55] + (
                  (z[56] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[56] + (
                  (z[57] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[57] + (
                  (z[58] - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c[58] + (
                  (z[59] - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c[59] + (max(c[0].value - 0, 0) * p1.VALUE) + (
                  max(c[12].value - c[0].value, 0) * p1.VALUE) + (
                  max(c[24].value - c[12].value, 0) * p1.VALUE) + (max(c[36].value - c[24].value, 0) * p1.VALUE)
          + (max(c[48].value - c[36].value, 0) * p1.VALUE) + (max(c[1].value - 0, 0) * p1.VALUE) + (
                  max(c[13].value - c[1].value, 0) * p1.VALUE) + (
                  max(c[25].value - c[13].value, 0) * p1.VALUE) + (max(c[37].value - c[25].value, 0) * p1.VALUE) + (
                  max(c[49].value - c[37].value, 0) * p1.VALUE) + (max(c[2].value - 0, 0) * p1.VALUE) + (
                  max(c[14].value - c[2].value, 0) * p1.VALUE) + (
                  max(c[26].value - c[14].value, 0) * p1.VALUE) + (max(c[38].value - c[26].value, 0) * p1.VALUE) + (
                  max(c[50].value - c[38].value, 0) * p1.VALUE) + (max(c[3].value - 0, 0) * p1.VALUE) + (
                  max(c[15].value - c[3].value, 0) * p1.VALUE) + (
                  max(c[27].value - c[15].value, 0) * p1.VALUE) + (max(c[39].value - c[27].value, 0) * p1.VALUE) + (
                  max(c[51].value - c[39].value, 0) * p1.VALUE) + (max(c[4].value - 0, 0) * p1.VALUE) + (
                  max(c[16].value - c[4].value, 0) * p1.VALUE) + (
                  max(c[28].value - c[16].value, 0) * p1.VALUE) + (max(c[40].value - c[28].value, 0) * p1.VALUE) + (
                  max(c[52].value - c[40].value, 0) * p1.VALUE) + (max(c[5].value - 0, 0) * p1.VALUE) + (
                  max(c[17].value - c[5].value, 0) * p1.VALUE) + (
                  max(c[29].value - c[17].value, 0) * p1.VALUE) + (max(c[41].value - c[29].value, 0) * p1.VALUE)
          + (max(c[53].value - c[41].value, 0) * p1.VALUE) + (max(c[6].value - 0, 0) * p1.VALUE) + (
                  max(c[18].value - c[6].value, 0) * p1.VALUE) + (
                  max(c[30].value - c[18].value, 0) * p1.VALUE) + (max(c[42].value - c[30].value, 0) * p1.VALUE) + (
                  max(c[54].value - c[42].value, 0) * p1.VALUE) + (max(c[7].value - 0, 0) * p1.VALUE) + (
                  max(c[19].value - c[7].value, 0) * p1.VALUE) + (
                  max(c[31].value - c[19].value, 0) * p1.VALUE) + (max(c[43].value - c[31].value, 0) * p1.VALUE) + (
                  max(c[55].value - c[43].value, 0) * p1.VALUE) + (max(c[8].value - 0, 0) * p1.VALUE) + (
                  max(c[20].value - c[18].value, 0) * p1.VALUE) + (
                  max(c[32].value - c[20].value, 0) * p1.VALUE) + (max(c[44].value - c[32].value, 0) * p1.VALUE) + (
                  max(c[56].value - c[44].value, 0) * p1.VALUE) + (max(c[9].value - 0, 0) * p1.VALUE) + (
                  max(c[21].value - c[9].value, 0) * p1.VALUE) + (
                  max(c[33].value - c[21].value, 0) * p1.VALUE) + (max(c[45].value - c[33].value, 0) * p1.VALUE) + (
                  max(c[57].value - c[45].value, 0) * p1.VALUE) + (max(c[10].value - 0, 0) * p1.VALUE) + (
                  max(c[22].value - c[10].value, 0) * p1.VALUE) + (
                  max(c[34].value - c[22].value, 0) * p1.VALUE) + (max(c[46].value - c[34].value, 0) * p1.VALUE) + (
                  max(c[58].value - c[46].value, 0) * p1.VALUE) + (max(c[11].value - 0, 0) * p1.VALUE) + (
                  max(c[23].value - c[11].value, 0) * p1.VALUE) + (
                  max(c[35].value - c[23].value, 0) * p1.VALUE) + (max(c[47].value - c[35].value, 0) * p1.VALUE) + (
                  max(c[59].value - c[47].value, 0) * p1.VALUE))

    m.options.IMODE = 3
    m.options.SOLVER = 1
    m.solver_options = ['minlp_maximum_iterations 10000',
                        'minlp_max_iter_with_int_sol 500',
                        'minlp_gap_tol 0.01',
                        'nlp_maximum_iterations 500',
                        'minlp_as_nlp 0',
                        'minlp_interger_leaves = 0',
                        'minlp_branch_method 2',
                        'minlp_integer_tol 0.01',
                        'minlp_print_level 2'
                        ]

    m.solve()

    k0 = np.array([k[i].value for i in range(25)])
    x0 = np.array([x[i].value for i in range(25)], dtype=str)

    for q in range(25):
        if k0[q] == 0:
            x0[q] = 'Выкл'

    j0 = np.array([j[i].value for i in range(50)])
    y0 = np.array([y[i].value for i in range(50)], dtype=str)

    for q1 in range(50):
        if j0[q1] == 0:
            y0[q1] = 'Выкл'

    c0 = np.array([c[i].value for i in range(60)])
    z0 = np.array([z[i].value for i in range(60)], dtype=str)

    for q2 in range(60):
        if c0[q2] == 0:
            z0[q2] = 'Выкл'

    df1 = pd.DataFrame(
        {'Период': ['c 00:00 до 06:00', 'с 06:00 до 09:00', 'с 09:00 до 15:00', 'с 15:00 до 18:00', 'с 18:00 до 00:00'],
         'Генератор №1 МВ': [z0[0], z0[12], z0[24], z0[36], z0[48]],
         'Генератор №2 МВ': [z0[1], z0[13], z0[25], z0[37], z0[49]],
         'Генератор №3 МВ': [z0[2], z0[14], z0[26], z0[38], z0[50]],
         'Генератор №4 МВ': [z0[3], z0[15], z0[27], z0[39], z0[51]],
         'Генератор №5 МВ': [z0[4], z0[16], z0[28], z0[40], z0[52]],
         'Генератор №6 МВ': [z0[5], z0[17], z0[29], z0[41], z0[53]],
         'Генератор №7 МВ': [z0[6], z0[18], z0[30], z0[42], z0[54]],
         'Генератор №8 МВ': [z0[7], z0[19], z0[31], z0[43], z0[55]],
         'Генератор №9 МВ': [z0[8], z0[20], z0[32], z0[44], z0[56]],
         'Генератор №10 МВ': [z0[9], z0[21], z0[33], z0[45], z0[57]],
         'Генератор №11 МВ': [z0[10], z0[22], z0[34], z0[46], z0[58]],
         'Генератор №12 МВ': [z0[11], z0[23], z0[35], z0[47], z0[59]]})

    df2 = pd.DataFrame(
        {'Период': ['c 00:00 до 06:00', 'с 06:00 до 09:00', 'с 09:00 до 15:00', 'с 15:00 до 18:00', 'с 18:00 до 00:00'],
         'Генератор №1 МВ': [y0[0], y0[10], y0[20], y0[30], y0[40]],
         'Генератор №2 МВ': [y0[1], y0[11], y0[21], y0[31], y0[41]],
         'Генератор №3 МВ': [y0[2], y0[12], y0[22], y0[33], y0[42]],
         'Генератор №4 МВ': [y0[3], y0[13], y0[23], y0[33], y0[43]],
         'Генератор №5 МВ': [y0[4], y0[14], y0[24], y0[34], y0[44]],
         'Генератор №6 МВ': [y0[5], y0[15], y0[25], y0[35], y0[45]],
         'Генератор №7 МВ': [y0[6], y0[16], y0[26], y0[36], y0[46]],
         'Генератор №8 МВ': [y0[7], y0[17], y0[27], y0[37], y0[47]],
         'Генератор №9 МВ': [y0[8], y0[18], y0[28], y0[38], y0[48]],
         'Генератор №10 МВ': [y0[9], y0[19], y0[29], y0[39], y0[49]]})

    df3 = pd.DataFrame(
        {'Период': ['c 00:00 до 06:00', 'с 06:00 до 09:00', 'с 09:00 до 15:00', 'с 15:00 до 18:00', 'с 18:00 до 00:00'],
         'Генератор №1 МВ': [x0[0], x0[5], x0[10], x0[15], x0[20]],
         'Генератор №2 МВ': [x0[1], x0[6], x0[11], x0[16], x0[21]],
         'Генератор №3 МВ': [x0[2], x0[7], x0[12], x0[17], x0[22]],
         'Генератор №4 МВ': [x0[3], x0[8], x0[13], x0[18], x0[23]],
         'Генератор №5 МВ': [x0[4], x0[9], x0[14], x0[19], x0[24]]})

    df4 = pd.DataFrame({'Минимальные издержки составляют:': [str(m.options.OBJFCNVAL)]})

    sheets = {'I тип': df1, 'II ТИП': df2, 'III ТИП': df3, 'Издержки': df4}
    writer = pd.ExcelWriter('./График_работы_генераторов.xlsx', engine='openpyxl')

    for sheet_name in sheets.keys():
        sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

    writer.save()
    open_button['state'] = 'normal'


root = Tk()
root.iconbitmap('icon.ico')
root.title("Оптимизация ТЭЦ")
root.geometry("600x300")

type1_label = Label(text="Стоимость запуска генератора I типа")
type2_label = Label(text="Стоимость запуска генератора II типа")
type3_label = Label(text="Стоимость запуска генератора III типа")

price1_label = Label(text="Стоимость одного часа работы генератора I типа на мин. мощности")
price2_label = Label(text="Стоимость одного часа работы генератора II типа на мин. мощности")
price3_label = Label(text="Стоимость одного часа работы генератора III типа на мин. мощности")

pricemw1_label = Label(text="Стоимость выработки одного МВт в час сверх мин. уровня генератором I типа")
pricemw2_label = Label(text="Стоимость выработки одного МВт в час сверх мин. уровня генератором II типа")
pricemw3_label = Label(text="Стоимость выработки одного МВт в час сверх мин. уровня генератором III типа")

type1_label.grid(row=0, column=0, sticky="w")
type2_label.grid(row=1, column=0, sticky="w")
type3_label.grid(row=2, column=0, sticky="w")
price1_label.grid(row=3, column=0, sticky="w")
price2_label.grid(row=4, column=0, sticky="w")
price3_label.grid(row=5, column=0, sticky="w")
pricemw1_label.grid(row=6, column=0, sticky="w")
pricemw2_label.grid(row=7, column=0, sticky="w")
pricemw3_label.grid(row=8, column=0, sticky="w")

type1_entry = Entry()
type2_entry = Entry()
type3_entry = Entry()
price1_entry = Entry()
price2_entry = Entry()
price3_entry = Entry()
pricemw1_entry = Entry()
pricemw2_entry = Entry()
pricemw3_entry = Entry()

type1_entry.grid(row=0, column=1, padx=5, pady=5)
type2_entry.grid(row=1, column=1, padx=5, pady=5)
type3_entry.grid(row=2, column=1, padx=5, pady=5)
price1_entry.grid(row=3, column=1, padx=5, pady=5)
price2_entry.grid(row=4, column=1, padx=5, pady=5)
price3_entry.grid(row=5, column=1, padx=5, pady=5)
pricemw1_entry.grid(row=6, column=1, padx=5, pady=5)
pricemw2_entry.grid(row=7, column=1, padx=5, pady=5)
pricemw3_entry.grid(row=8, column=1, padx=5, pady=5)

type1_entry.insert(0, "150000")
type2_entry.insert(0, "60000")
type3_entry.insert(0, "90000")

price1_entry.insert(0, "175000")
price2_entry.insert(0, "200000")
price3_entry.insert(0, "390000")

pricemw1_entry.insert(0, "150")
pricemw2_entry.insert(0, "120")
pricemw3_entry.insert(0, "200")

start_button = Button(text="Запустить", command=mat)
open_button = Button(text="Открыть excel", command=openfile)

start_button.grid(row=9, column=1, padx=50, pady=5, sticky="e")
open_button.grid(row=9, column=0, padx=5, pady=5, sticky="w")
open_button['state'] = 'disabled'

root.mainloop()
