GPR = "GPR地质预报"
TSP = "TSP地质预报"
PS = "掌子面地质素描"
CP = "施工进度"
CA = "施工变更"

S1S2 = "S1S2"
S3S4 = "S3S4"

tasks = {
    # "GPRs1": ["GPR.FileProcess_GPR_S1S2", "FileProcess_GPR_S1S2"],
    # "GPRs3": ["S3S4.FileProcess_GPR_S3S4", "FileProcess_GPR_S3S4"],
    # (GPR, S1S2): "library.GPR.FileProcess_GPR_S1S2",
    (PS, S1S2): "library.PS.FileProcess_PS_S1S2",
    # (TSP, S1S2): "library.TSP.FileProcess_TSP_S1S2",

    # (GPR, S3S4): "library.GPR.FileProcess_GPR_S3S4",
    # (CA, S3S4): "library.CHAG.FileProcess_CHAG",
    # (CA, S1S2): "library.CHAG.FileProcess_CHAG",
    # (CA, None): "library.CHAG.FileProcess_CHAG",
}

projects = {
    GPR: True,
    TSP: True,
    PS: True,
    CP: False,
    CA: True
}

datatypes = {
    "s1": S1S2,
    "s2": S1S2,
    "s3": S3S4,
    "s4": S3S4
}
