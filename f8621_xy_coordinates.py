coordinate_dict = {
    # persoanl info
    "Name of shareholder": (36, 688),  # Name
    "Identifying Number": (336, 688),  # SS Number
    "Address": (36, 662),  # Address
    "City, State, Zip, Country": (36, 638),  # city stare zip
    "Tax year": (461, 675),  # tax year
    "Type of Shareholder": (196, 627),  # type of shareholder
    # PFIC info
    "PFIC Name": (36, 567),  # PFIC
    "PFIC Address": (36, 543),  # PFIC address
    "PFIC Reference ID": (361, 543),  # PFIC ref id
    "PFIC Share Class": (281, 470),  # Descrition of each class of shares
    # PART I
    "Date of Acquisition": (263, 434),  # Date of Acquisition
    "Number of Shares": (243, 410),  # number of shares
    "Amount of 1291": (152, 314),  # amount of 1291
    "Amount of 1293": (245.6, 302),  # amount of 1293
    "Amount of 1296": (217, 290),  # amount of 1296
    "Amount of 1296- Check": (79.2, 290),  # type of PFIC type c
    # Part II
    "Check MTM": (52.4, 205.5),
    # PART IV
    "10a": (489.606, 408.01),
    "10b": (489.606, 396.011),
    "10c": (489.606, 372.007),
    "11": (489.606, 360.008),
    "12": (489.606, 336.007),
    "13a": (489.606, 312.009),
    "13b": (489.606, 300.007),
    "13c": (489.606, 276.009),
    "14a": (489.606, 264.01),
    "14b": (489.606, 228.01),
    "14c": (489.606, 192.01),
}


def get_coordinates():
    return coordinate_dict
