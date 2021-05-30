from gekko import GEKKO
import numpy as np
from tkinter import *
import pandas as pd
import openpyxl
import os


def openfile():
    os.chdir('C:\\Users\\evilm\\PycharmProjects\\pythonProject')
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

    x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18, x19, x20, x21, x22, x23, x24, x25 = \
        [m.Var(value=1500, lb=1500, ub=4000) for i in range(25)]
    k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11, k12, k13, k14, k15, k16, k17, k18, k19, k20, k21, k22, k23, k24, k25 = \
        [m.Var(value=1, lb=0, ub=1, integer=True) for i1 in range(25)]

    y1, y2, y3, y4, y5, y6, y7, y8, y9, y10, y11, y12, y13, y14, y15, y16, y17, y18, y19, y20, y21, y22, y23, y24, y25, \
    y26, y27, y28, y29, y30, y31, y32, y33, y34, y35, y36, y37, y38, y39, y40, y41, y42, y43, y44, y45, y46, y47, y48, \
    y49, y50 = [m.Var(value=1250, lb=1250, ub=1750, ) for i2 in range(50)]
    j1, j2, j3, j5, j4, j6, j7, j8, j9, j10, j11, j12, j13, j14, j15, j16, j17, j18, j19, j20, j21, j22, j23, j24, j25, \
    j26, j27, j28, j29, j30, j31, j32, j33, j34, j35, j36, j37, j38, j39, j40, j41, j42, j43, j44, j45, j46, j47, j48, \
    j49, j50 = [m.Var(value=1, lb=0, ub=1, integer=True) for i3 in range(50)]

    z1, z2, z3, z4, z5, z6, z7, z8, z9, z10, z11, z12, z13, z14, z15, z16, z17, z18, z19, z20, z21, z22, z23, z24, z25, \
    z26, z27, z28, z29, z30, z31, z32, z33, z34, z35, z36, z37, z38, z39, z40, z41, z42, z43, z44, z45, z46, z47, z48, \
    z49, z50, z51, z52, z53, z54, z55, z56, z57, z58, z59, z60 = [m.Var(value=850, lb=850, ub=2000, ) for i4 in
                                                                  range(60)]
    c1, c2, c3, c5, c4, c6, c7, c8, c9, c10, c11, c12, c13, c14, c15, c16, c17, c18, c19, c20, c21, c22, c23, c24, c25, \
    c26, c27, c28, c29, c30, c31, c32, c33, c34, c35, c36, c37, c38, c39, c40, c41, c42, c43, c44, c45, c46, c47, c48, c49, \
    c50, c51, c52, c53, c54, c55, c56, c57, c58, c59, c60 = [m.Var(value=1, lb=0, ub=1, integer=True) for i5 in
                                                             range(60)]

    m.Equation(
        x1 * k1 + x2 * k2 + x3 * k3 + x4 * k4 + x5 * k5 + y1 * j1 + y2 * j2 + y3 * j3 + y4 * j4 + y5 * j5 + y6 * j6 + y7 * j7 + y8 * j8 + y9 * j9 + y10 * j10 + z1 * c1 + z2 * c2 + z3 * c3 + z4 * c4 + z5 * c5 + z6 * c6 + z7 * c7 + z8 * c8 + z9 * c9 + z10 * c10 + z11 * c11 + z12 * c12 >= 17250)
    m.Equation(
        x6 * k6 + x7 * k7 + x8 * k8 + x9 * k9 + x10 * k10 + y11 * j11 + y12 * j12 + y13 * j13 + y14 * j14 + y15 * j15 + y16 * j16 + y17 * j17 + y18 * j18 + y19 * j19 + y20 * j20 + z13 * c13 + z14 * c14 + z15 * c15 + z16 * c16 + z17 * c17 + z18 * c18 + z19 * c19 + z20 * c20 + z21 * c21 + z22 * c22 + z23 * c23 + z24 * c24 >= 34500)
    m.Equation(
        x11 * k11 + x12 * k12 + x13 * k13 + x14 * k14 + x15 * k15 + y21 * j21 + y22 * j22 + y23 * j23 + y24 * j24 + y25 * j25 + y26 * j26 + y27 * j27 + y28 * j28 + y29 * j29 + y30 * j30 + z25 * c25 + z26 * c26 + z27 * c27 + z28 * c28 + z29 * c29 + z30 * c30 + z31 * c31 + z32 * c32 + z33 * c33 + z34 * c34 + z35 * c35 + z36 * c36 >= 28750)
    m.Equation(
        x16 * k16 + x17 * k17 + x18 * k18 + x19 * k19 + x20 * k20 + y31 * j31 + y32 * j32 + y33 * j33 + y34 * j34 + y35 * j35 + y36 * j36 + y37 * j37 + y38 * j38 + y39 * j39 + y40 * j40 + z37 * c37 + z38 * c38 + z39 * c39 + z40 * c40 + z41 * c41 + z42 * c42 + z43 * c43 + z44 * c44 + z45 * c45 + z46 * c46 + z47 * c47 + z48 * c48 >= 46000)
    m.Equation(
        x21 * k21 + x22 * k22 + x23 * k23 + x24 * k24 + x25 * k25 + y41 * j41 + y42 * j42 + y43 * j43 + y44 * j44 + y45 * j45 + y46 * j46 + y47 * j47 + y48 * j48 + y49 * j49 + y50 * j50 + z49 * c49 + z50 * c50 + z51 * c51 + z52 * c52 + z53 * c53 + z54 * c54 + z55 * c55 + z56 * c56 + z57 * c57 + z58 * c58 + z59 * c59 + z60 * c60 >= 31050)

    p1 = m.Param(int(price_for_start1))
    p2 = m.Param(int(price_for_start2))
    p3 = m.Param(int(price_for_start3))

    ph1 = m.Param(int(price_for_hour1))
    ph2 = m.Param(int(price_for_hour2))
    ph3 = m.Param(int(price_for_hour3))

    pm3 = m.Param(int(price_for_mw3))
    pm2 = m.Param(int(price_for_mw2))
    pm1 = m.Param(int(price_for_mw1))

    m.Obj(((x1 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k1 + ((x2 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k2 + (
            (x3 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k3 + (
                  (x4 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k4 + (
                  (x5 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k5 + (
                  (x6 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k6 + (
                  (x7 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k7 + (
                  (x8 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k8 + (
                  (x9 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k9 + (
                  (x10 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k10 + (
                  (x11 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k11 + (
                  (x12 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k12 + (
                  (x13 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k13 + (
                  (x14 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k14 + (
                  (x15 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k15 + (
                  (x16 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k16 + (
                  (x17 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k17 + (
                  (x18 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k18 + (
                  (x19 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k19 + (
                  (x20 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k20 + (
                  (x21 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k21 + (
                  (x22 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k22 + (
                  (x23 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k23 + (
                  (x24 - 1500) * pm3.VALUE * 3 + ph3.VALUE * 3) * k24 + (
                  (x25 - 1500) * pm3.VALUE * 6 + ph3.VALUE * 6) * k25 + (max(k1.value - 0, 0) * p3.VALUE) + (
                  max(k6.value - k1.value, 0) * p3.VALUE) + (
                  max(k11.value - k6.value, 0) * p3.VALUE) + (max(k16.value - k11.value, 0) * p3.VALUE)
          + (max(k21.value - k16.value, 0) * p3.VALUE) + (max(k2.value - 0, 0) * p3.VALUE) + (
                  max(k7.value - k2.value, 0) * p3.VALUE) + (
                  max(k12.value - k7.value, 0) * p3.VALUE) + (max(k17.value - k12.value, 0) * p3.VALUE) + (
                  max(k22.value - k17.value, 0) * p3.VALUE) + (max(k3.value - 0, 0) * p3.VALUE) + (
                  max(k8.value - k3.value, 0) * p3.VALUE) + (
                  max(k13.value - k8.value, 0) * p3.VALUE) + (max(k18.value - k13.value, 0) * p3.VALUE) + (
                  max(k23.value - k18.value, 0) * p3.VALUE) + (max(k4.value - 0, 0) * p3.VALUE) + (
                  max(k9.value - k4.value, 0) * p3.VALUE) + (
                  max(k14.value - k9.value, 0) * p3.VALUE) + (max(k19.value - k14.value, 0) * p3.VALUE) + (
                  max(k24.value - k19.value, 0) * p3.VALUE) + (max(k5.value - 0, 0) * p3.VALUE) + (
                  max(k10.value - k5.value, 0) * p3.VALUE) + (
                  max(k15.value - k10.value, 0) * p3.VALUE) + (max(k20.value - k15.value, 0) * p3.VALUE) + (
                  max(k25.value - k20.value, 0) * p3.VALUE) + ((y1 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j1 + (
                  (y2 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j2 + (
                  (y3 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j3 + (
                  (y4 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j4 + (
                  (y5 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j5 + (
                  (y6 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j6 + (
                  (y7 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j7 + (
                  (y8 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j8 + (
                  (y9 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j9 + (
                  (y10 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j10 + (
                  (y11 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j11 + (
                  (y12 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j12 + (
                  (y13 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j13 + (
                  (y14 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j14 + (
                  (y15 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j15 + (
                  (y16 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j16 + (
                  (y17 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j17 + (
                  (y18 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j18 + (
                  (y19 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j19 + (
                  (y20 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j20 + (
                  (y21 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j21 + (
                  (y22 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j22 + (
                  (y23 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j23 + (
                  (y24 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j24 + (
                  (y25 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j25 + (
                  (y26 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j26 + (
                  (y27 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j27 + (
                  (y28 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j28 + (
                  (y29 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j29 + (
                  (y30 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j30 + (
                  (y31 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j31 + (
                  (y32 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j32 + (
                  (y33 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j33 + (
                  (y34 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j34 + (
                  (y35 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j35 + (
                  (y36 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j36 + (
                  (y37 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j37 + (
                  (y38 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j38 + (
                  (y39 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j39 + (
                  (y40 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j40 + (
                  (y41 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j41 + (
                  (y42 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j42 + (
                  (y43 - 1250) * pm2.VALUE * 6 + ph2.VALUE) * j43 + (
                  (y44 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j44 + (
                  (y45 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j45 + (
                  (y46 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j46 + (
                  (y47 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j47 + (
                  (y48 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j48 + (
                  (y49 - 1250) * pm2.VALUE * 6 + ph2.VALUE * 6) * j49 + (
                  (y50 - 1250) * pm2.VALUE * 3 + ph2.VALUE * 3) * j50 + (max(j1.value - 0, 0) * p2.VALUE) + (
                  max(j11.value - j1.value, 0) * p2.VALUE) + (
                  max(j21.value - j11.value, 0) * p2.VALUE) + (max(j31.value - j21.value, 0) * p2.VALUE)
          + (max(j41.value - j31.value, 0) * p2.VALUE) + (max(j2.value - 0, 0) * p2.VALUE) + (
                  max(j12.value - j2.value, 0) * p2.VALUE) + (
                  max(j22.value - j12.value, 0) * p2.VALUE) + (max(j32.value - j22.value, 0) * p2.VALUE) + (
                  max(j42.value - j32.value, 0) * p2.VALUE) + (max(j3.value - 0, 0) * p2.VALUE) + (
                  max(j13.value - j3.value, 0) * p2.VALUE) + (
                  max(j23.value - j13.value, 0) * p2.VALUE) + (max(j33.value - j23.value, 0) * p2.VALUE) + (
                  max(j43.value - j23.value, 0) * p2.VALUE) + (max(j4.value - 0, 0) * p2.VALUE) + (
                  max(j14.value - j4.value, 0) * p2.VALUE) + (
                  max(j24.value - j14.value, 0) * p2.VALUE) + (max(j34.value - j24.value, 0) * p2.VALUE) + (
                  max(j44.value - j34.value, 0) * p2.VALUE) + (max(j5.value - 0, 0) * p2.VALUE) + (
                  max(j15.value - j5.value, 0) * p2.VALUE) + (
                  max(j25.value - j15.value, 0) * p2.VALUE) + (max(j35.value - j25.value, 0) * p2.VALUE) + (
                  max(j45.value - j35.value, 0) * p2.VALUE) + (max(j6.value - 0, 0) * p2.VALUE) + (
                  max(j16.value - j6.value, 0) * p2.VALUE) + (
                  max(j26.value - j16.value, 0) * p2.VALUE) + (max(j36.value - j26.value, 0) * p2.VALUE)
          + (max(j46.value - j36.value, 0) * p2.VALUE) + (max(j7.value - 0, 0) * p2.VALUE) + (
                  max(j17.value - j7.value, 0) * p2.VALUE) + (
                  max(j27.value - j17.value, 0) * p2.VALUE) + (max(j37.value - j27.value, 0) * p2.VALUE) + (
                  max(j47.value - j37.value, 0) * p2.VALUE) + (max(j8.value - 0, 0) * p2.VALUE) + (
                  max(j18.value - j8.value, 0) * p2.VALUE) + (
                  max(j28.value - j18.value, 0) * p2.VALUE) + (max(j38.value - j28.value, 0) * p2.VALUE) + (
                  max(j48.value - j38.value, 0) * p2.VALUE) + (max(j9.value - 0, 0) * p2.VALUE) + (
                  max(j19.value - j9.value, 0) * p2.VALUE) + (
                  max(j29.value - j19.value, 0) * p2.VALUE) + (max(j39.value - j29.value, 0) * p2.VALUE) + (
                  max(j49.value - j39.value, 0) * p2.VALUE) + (max(j10.value - 0, 0) * p2.VALUE) + (
                  max(j20.value - j10.value, 0) * p2.VALUE) + (
                  max(j30.value - j20.value, 0) * p2.VALUE) + (max(j40.value - j30.value, 0) * p2.VALUE) + (
                  max(j50.value - j40.value, 0) * p2.VALUE) + ((z1 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c1 + (
                  (z2 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c2 + (
                  (z3 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c3 + (
                  (z4 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c4 + (
                  (z5 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c5 + (
                  (z6 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c6 + (
                  (z7 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c7 + (
                  (z8 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c8 + (
                  (z9 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c9 + (
                  (z10 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c10 + (
                  (z11 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c11 + (
                  (z12 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c12 + (
                  (z13 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c13 + (
                  (z14 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c14 + (
                  (z15 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c15 + (
                  (z16 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c16 + (
                  (z17 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c17 + (
                  (z18 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c18 + (
                  (z19 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c19 + (
                  (z20 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c20 + (
                  (z21 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c21 + (
                  (z22 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c22 + (
                  (z23 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c23 + (
                  (z24 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c24 + (
                  (z25 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c25 + (
                  (z26 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c26 + (
                  (z27 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c27 + (
                  (z28 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c28 + (
                  (z29 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c29 + (
                  (z30 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c30 + (
                  (z31 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c31 + (
                  (z32 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c32 + (
                  (z33 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c33 + (
                  (z34 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c34 + (
                  (z35 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c35 + (
                  (z36 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c36 + (
                  (z37 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c37 + (
                  (z38 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c38 + (
                  (z39 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c39 + (
                  (z40 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c40 + (
                  (z41 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c41 + (
                  (z42 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c42 + (
                  (z43 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c43 + (
                  (z44 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c44 + (
                  (z45 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c45 + (
                  (z46 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c46 + (
                  (z47 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c47 + (
                  (z48 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c48 + (
                  (z49 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c49 + (
                  (z50 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c50 + (
                  (z51 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c51 + (
                  (z52 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c52 + (
                  (z53 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c53 + (
                  (z54 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c54 + (
                  (z55 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c55 + (
                  (z56 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c56 + (
                  (z57 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c57 + (
                  (z58 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c58 + (
                  (z59 - 850) * pm1.VALUE * 6 + ph1.VALUE * 6) * c59 + (
                  (z60 - 850) * pm1.VALUE * 3 + ph1.VALUE * 3) * c60 + (max(c1.value - 0, 0) * p1.VALUE) + (
                  max(c13.value - c1.value, 0) * p1.VALUE) + (
                  max(c25.value - c13.value, 0) * p1.VALUE) + (max(j37.value - c25.value, 0) * p1.VALUE)
          + (max(c49.value - c37.value, 0) * p1.VALUE) + (max(c2.value - 0, 0) * p1.VALUE) + (
                  max(c14.value - c2.value, 0) * p1.VALUE) + (
                  max(c26.value - c14.value, 0) * p1.VALUE) + (max(c38.value - c26.value, 0) * p1.VALUE) + (
                  max(c50.value - c38.value, 0) * p1.VALUE) + (max(c3.value - 0, 0) * p1.VALUE) + (
                  max(c15.value - c3.value, 0) * p1.VALUE) + (
                  max(c27.value - c15.value, 0) * p1.VALUE) + (max(c39.value - c27.value, 0) * p1.VALUE) + (
                  max(c51.value - c39.value, 0) * p1.VALUE) + (max(c4.value - 0, 0) * p1.VALUE) + (
                  max(c16.value - c4.value, 0) * p1.VALUE) + (
                  max(c28.value - c16.value, 0) * p1.VALUE) + (max(c40.value - c28.value, 0) * p1.VALUE) + (
                  max(c52.value - c40.value, 0) * p1.VALUE) + (max(c5.value - 0, 0) * p1.VALUE) + (
                  max(c17.value - c5.value, 0) * p1.VALUE) + (
                  max(c29.value - c17.value, 0) * p1.VALUE) + (max(c41.value - c29.value, 0) * p1.VALUE) + (
                  max(c53.value - c41.value, 0) * p1.VALUE) + (max(c6.value - 0, 0) * p1.VALUE) + (
                  max(c18.value - c6.value, 0) * p1.VALUE) + (
                  max(c30.value - c18.value, 0) * p1.VALUE) + (max(c42.value - c30.value, 0) * p1.VALUE)
          + (max(c54.value - c42.value, 0) * p1.VALUE) + (max(c7.value - 0, 0) * p1.VALUE) + (
                  max(c19.value - j7.value, 0) * p1.VALUE) + (
                  max(c31.value - c19.value, 0) * p1.VALUE) + (max(c43.value - c31.value, 0) * p1.VALUE) + (
                  max(c55.value - c43.value, 0) * p1.VALUE) + (max(c8.value - 0, 0) * p1.VALUE) + (
                  max(c20.value - c8.value, 0) * p1.VALUE) + (
                  max(c32.value - c20.value, 0) * p1.VALUE) + (max(c44.value - c32.value, 0) * p1.VALUE) + (
                  max(c56.value - c44.value, 0) * p1.VALUE) + (max(c9.value - 0, 0) * p1.VALUE) + (
                  max(c21.value - c19.value, 0) * p1.VALUE) + (
                  max(c33.value - c21.value, 0) * p1.VALUE) + (max(c45.value - c33.value, 0) * p1.VALUE) + (
                  max(c57.value - c45.value, 0) * p1.VALUE) + (max(c10.value - 0, 0) * p1.VALUE) + (
                  max(c22.value - c10.value, 0) * p1.VALUE) + (
                  max(c34.value - c22.value, 0) * p1.VALUE) + (max(c46.value - c34.value, 0) * p1.VALUE) + (
                  max(c58.value - c46.value, 0) * p1.VALUE) + (max(c11.value - 0, 0) * p1.VALUE) + (
                  max(c23.value - c11.value, 0) * p1.VALUE) + (
                  max(c35.value - c23.value, 0) * p1.VALUE) + (max(c47.value - c35.value, 0) * p1.VALUE) + (
                  max(c59.value - c47.value, 0) * p1.VALUE) + (max(c12.value - 0, 0) * p1.VALUE) + (
                  max(c24.value - c12.value, 0) * p1.VALUE) + (
                  max(c36.value - c24.value, 0) * p1.VALUE) + (max(c48.value - c36.value, 0) * p1.VALUE) + (
                  max(c60.value - c48.value, 0) * p1.VALUE))

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

    k0 = np.array(
        [k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11, k12, k13, k14, k15, k16, k17, k18, k19, k20, k21, k22, k23, k24,
         k25])
    x0 = np.array(
        [x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18, x19, x20, x21, x22, x23, x24,
         x25], dtype=str)

    for q in range(25):
        if k0[q] == 0:
            x0[q] = 'Выкл'

    j0 = np.array(
        [j1, j2, j3, j5, j4, j6, j7, j8, j9, j10, j11, j12, j13, j14, j15, j16, j17, j18, j19, j20, j21, j22, j23, j24,
         j25,
         j26, j27, j28, j29, j30, j31, j32, j33, j34, j35, j36, j37, j38, j39, j40, j41, j42, j43, j44, j45, j46, j47,
         j48,
         j49, j50])
    y0 = np.array(
        [y1, y2, y3, y4, y5, y6, y7, y8, y9, y10, y11, y12, y13, y14, y15, y16, y17, y18, y19, y20, y21, y22, y23, y24,
         y25,
         y26, y27, y28, y29, y30, y31, y32, y33, y34, y35, y36, y37, y38, y39, y40, y41, y42, y43, y44, y45, y46, y47,
         y48,
         y49, y50], dtype=str)

    for q1 in range(50):
        if j0[q1] == 0:
            y0[q1] = 'Выкл'

    c0 = np.array(
        [c1, c2, c3, c5, c4, c6, c7, c8, c9, c10, c11, c12, c13, c14, c15, c16, c17, c18, c19, c20, c21, c22, c23, c24,
         c25,
         c26, c27, c28, c29, c30, c31, c32, c33, c34, c35, c36, c37, c38, c39, c40, c41, c42, c43, c44, c45, c46, c47,
         c48,
         c49, c50, c51, c52, c53, c54, c55, c56, c57, c58, c59, c60])
    z0 = np.array(
        [z1, z2, z3, z4, z5, z6, z7, z8, z9, z10, z11, z12, z13, z14, z15, z16, z17, z18, z19, z20, z21, z22, z23, z24,
         z25,
         z26, z27, z28, z29, z30, z31, z32, z33, z34, z35, z36, z37, z38, z39, z40, z41, z42, z43, z44, z45, z46, z47,
         z48,
         z49, z50, z51, z52, z53, z54, z55, z56, z57, z58, z59, z60], dtype=str)

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
