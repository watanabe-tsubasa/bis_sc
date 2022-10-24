list_17 = list(range(201703, 201713)) + list(range(201801, 201803))
list_18 = list(range(201803, 201813)) + list(range(201901, 201903))
list_19 = list(range(201903, 201913)) + list(range(202001, 202003))
list_20 = list(range(202003, 202013)) + list(range(202101, 202103))
list_21 = list(range(202103, 202113)) + list(range(202201, 202203))
list_22 = list(range(202203, 202210)) # + list(range(202201, 202203))

years = {
    2017: list_17,
    2018: list_18,
    2019: list_19,
    2020: list_20,
    2021: list_21,
    2022: list_22
}


for i,j in years.items():
    print(j)