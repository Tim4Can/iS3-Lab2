GPR = "GPR地质预报"
TSP = "TSP地质预报"
PS = "掌子面地质素描"
CP = "施工进度"
CA = "施工变更"

S1S2 = "S1S2"
S3S4 = "S3S4"

datatypes = {
    "s1": S1S2,
    "s2": S1S2,
    "s3": S3S4,
    "s4": S3S4
}

tasks = {

   # (TSP, S1S2): "script.TSP.test_TSP_S1S2",
    #(CA, S3S4): "script.CHAG.test_CHAG_S3S4",
    (TSP, S3S4): "script.TSP.test_TSP_S3S4",

}

projects = {
    GPR: True,
    TSP: True,
    PS: True,
    CP: False,
    CA: True
}