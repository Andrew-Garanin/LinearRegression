import numpy as np
import scipy.stats as st
import matplotlib.pyplot as plt
import openpyxl


def main():
    wb = openpyxl.load_workbook(filename='data/lengthAndWidth.xlsx')
    sheets = wb.sheetnames
    sheet = wb[sheets[0]]
    x = [c[0].value for c in sheet['A1':'A20']]
    y = [c[0].value for c in sheet['B1':'B20']]

    # Выборочные характеристики
    mx = np.mean(x)
    my = np.mean(y)
    sx = np.std(x)
    sy = np.std(y)
    print("M(X) = %.3f, M(Y) = %.3f" % (mx, my))
    print("s(X) = %.3f, s(Y) = %.3f" % (sx, sy))

    # Выборочные коэф. коррел.
    print("H_0={r(X, Y) = 0}")

    r = st.pearsonr(x, y)[0]
    print("r = %.3f" % r)

    t_krit = 2.1009  # k=18, alpha = 0.05
    t_nabl = 18 ** 0.5 * r / (1 - r ** 2) ** 0.5

    print("T_krit = %.3f" % t_krit)
    print("T_nabl = %.3f" % t_nabl)

    if abs(t_nabl) > t_krit:
        print("H_0 отвергается. Между величинами обнаружена значимая корреляционная связь")
    else:
        print("Нет оснований отвернуть гипотезу H_0.")
        return

    # Коэфф. регрессии y = ax+b
    a = r * sy / sx
    b = my - r * sy * mx / sx
    print("a = %.3f, b = %.3f" % (a, b))
    if b >= 0:
        print("Уравнение линейной регрессии имеет вид: y = %.3fx+%.3f" % (a, b))
    else:
        print("Уравнение линейной регрессии имеет вид: y = %.3fx%.3f" % (a, b))

    # Проверка значимости
    sum = 0
    for i, j in zip(x, y):
        sum += (j - a * i - b) ** 2
    s_ost = (sum / (20 - 2)) ** 0.5
    s_a = s_ost / (sx * (20 ** 0.5))

    x_square = [xi ** 2 for xi in x]

    s_b = s_ost / (sx * (20 ** 0.5)) * np.mean(x_square)
    t_a = a / s_a
    t_b = b / s_b

    if abs(t_a) < t_krit:
        print("Коэф. a незначим => отказываемся от линейной модели")
    elif abs(t_b) < t_krit:
        print("Коэф. b незначим => уравнение линейной регрессии имеет вид: y = %.3fx" % a)
    else:
        print("Все коэффициенты значимы")

    # Прямые регрессий
    fig, ax = plt.subplots()
    plt.scatter(x, y)

    [xl, yl] = [[min(x), max(x)], [a * min(x) + b, a * max(x) + b]]
    ax.plot(xl, yl, label='С учётом коэффициента b')

    [xl_new, yl_new] = [[min(x), max(x)], [a * min(x), a * max(x)]]
    ax.plot(xl_new, yl_new, label='Без учёта коэффициента b')

    legend = ax.legend()
    legend.get_frame().set_facecolor('C0')
    plt.show()


if __name__ == '__main__':
    main()

