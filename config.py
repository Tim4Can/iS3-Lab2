'''config.py'''

GPR = "GPR地质预报"
TSP = "TSP地质预报"
PS = "掌子面地质素描"
CP = "施工进度"
CA = "施工变更"

S1S2 = "S1S2"
S3S4 = "S3S4"

tasks = {
    # "GPRs1": ["S1S2.FileProcess_GPR_S1S2", "FileProcess_GPR_S1S2"],
    # "GPRs3": ["S3S4.FileProcess_GPR_S3S4", "FileProcess_GPR_S3S4"],
    (GPR, S1S2): "library.S1S2.FileProcess_GPR_S1S2",
    # (GPR, S3S4): "S3S4\\FileProcess_GPR_S3S4"
}

projects = {
    GPR: True,
    TSP: True,
    PS: True,
    CP: False,
    CA: False
}

datatypes = {
    "s1": S1S2,
    "s2": S1S2,
    "s3": S3S4,
    "s4": S3S4
}
