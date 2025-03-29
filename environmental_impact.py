import brightway2 as bw
import pandas as pd
import matplotlib.pyplot as plt

bw.projects.set_current("EnvironmentalRiskProject")
bw.bw2setup()

db = bw.Database("simple_db")
db.write({
    ("simple_db", "ChemicalSpill"): {
        "name": "Chemical Spill Impact",
        "unit": "unit",
        "exchanges": [
            {"input": ("simple_db", "CO2"), "amount": 100, "type": "biosphere"},
            {"input": ("simple_db", "SO2"), "amount": 20, "type": "biosphere"},
        ],
    },
    ("simple_db", "WaterContamination"): {
        "name": "Water Contamination Impact",
        "unit": "unit",
        "exchanges": [
            {"input": ("simple_db", "Pollutant"), "amount": 50, "type": "biosphere"},
            {"input": ("simple_db", "SO2"), "amount": 10, "type": "biosphere"},
        ],
    },
    ("simple_db", "EcosystemDamage"): {
        "name": "Ecosystem Damage Impact",
        "unit": "unit",
        "exchanges": [
            {"input": ("simple_db", "CO2"), "amount": 30, "type": "biosphere"},
            {"input": ("simple_db", "SO2"), "amount": 5, "type": "biosphere"},
        ],
    },
    ("simple_db", "CO2"): {
        "name": "CO2 Emission",
        "type": "biosphere",
        "unit": "kg",
    },
    ("simple_db", "Pollutant"): {
        "name": "Water Pollutant",
        "type": "biosphere",
        "unit": "kg",
    },
    ("simple_db", "SO2"): {
        "name": "SO2 Emission",
        "type": "biosphere",
        "unit": "kg",
    },
})

gwp_method = ("IPCC", "climate change", "GWP 100a")
acid_method = ("CML", "acidification", "generic")

bw.Method(gwp_method).write([
    (db.get("CO2").key, {"amount": 1}),
    (db.get("Pollutant").key, {"amount": 2}),
    (db.get("SO2").key, {"amount": 0}),
])

bw.Method(acid_method).write([
    (db.get("CO2").key, {"amount": 0}),
    (db.get("Pollutant").key, {"amount": 0.5}),
    (db.get("SO2").key, {"amount": 1.5}),
])

results_gwp = []
results_acid = []
processes = [db.get("ChemicalSpill"), db.get("WaterContamination"), db.get("EcosystemDamage")]

for process in processes:
    lca = bw.LCA({process: 1}, gwp_method)
    lca.lci()
    lca.lcia()
    results_gwp.append({"Risk Name": process["name"].replace(" Impact", ""), "GWP (kg CO2 eq)": lca.score})

    lca.switch_method(acid_method)
    lca.lcia()
    results_acid.append({"Risk Name": process["name"].replace(" Impact", ""), "Acidification (kg SO2 eq)": lca.score})

df_gwp = pd.DataFrame(results_gwp)
df_acid = pd.DataFrame(results_acid)
df = df_gwp.merge(df_acid, on="Risk Name")

print("Environmental Impact Results:")
print(df.to_string(index=False))

plt.figure(figsize=(10, 6))
plt.bar(df["Risk Name"], df["GWP (kg CO2 eq)"], color="blue", label="GWP (kg CO2 eq)")
plt.bar(df["Risk Name"], df["Acidification (kg SO2 eq)"], color="red", alpha=0.5, label="Acidification (kg SO2 eq)")
plt.xlabel("Risk Name")
plt.ylabel("Impact")
plt.title("Environmental Impacts (GWP and Acidification)")
plt.legend()
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
plt.savefig("environmental_impact_bar.png")
plt.close()

print("Chart saved as 'environmental_impact_bar.png'")