import argparse
import csv
import os
import psr.graf
import powerfactory


_HAS_PANDAS = False
try:
    import pandas as pd
    _HAS_PANDAS = True
except ImportError:
    pd = None

_HAS_GRAF = False
try:
    import psr.graf
    _HAS_GRAF = True
except ImportError:
    psr = None
    pass

_DEBUG_PRINT = False


_PLANT_TYPE_OUTPUT_MAP = {
    "hydro": "gerhid",
    "thermal": "gerter",
    "renewable": "gergnd",
    "battery": "gerbat",
    "csp": "cspgen",
    "injection": "powinj",
}

_DURATION_FILE = "duraci"

_ELM_CLASS_ATTRIBUTE_MAP = {
    "ElmSym": "pgini",
    "ElmAsm": "pgini",
    "ElmGenstat": "pgini",
    "ElmAsmsc": "pgini",
    "ElmPvsys": "pgini",
    "ElmXnet": "pgini",
}

_ELM_CLASS_ATTRIBUTE_FACTOR_MAP = {
    "ElmSym": 1.0,  # MW
    "ElmAsm": 1.0,  # MW
    "ElmGenstat": 1.0,  # MW
    "ElmAsmsc": 1.0,  # MW
    "ElmPvsys": 0.001,  # kV
    "ElmXnet": 1.0,  # MW
}


class PlantMapEntry:
    def __init__(self):
        self.plant = SddpPlant()
        self.weight = 0.0
        self.elm_name = ""


class SddpPlant:
    def __init__(self):
        self.system = ""
        self.name = ""
        self.type = ""

    def __hash__(self):
        return hash("{},{},{}".format(self.system, self.name, self.type))

    def __eq__(self, other):
        return hash(self) == hash(other)


class SddpScenario:
    def __init__(self):
        self.stage = 0
        self.scenario = 0
        self.block = 0

    def __hash__(self):
        return hash("{},{},{}".format(self.stage, self.scenario, self.block))

    def __eq__(self, other):
        return hash(self) == hash(other)

    def as_tuple(self):
        return self.stage, self.scenario, self.block


def create_scenario(app, scenario_name: str):
    app.PrintInfo("Creating scenario: " + scenario_name)
    scen_folder = app.GetProjectFolder('scen')
    scenario = scen_folder.CreateObject("IntScenario", scenario_name)
    return scenario


def _read_plant_map(plant_map_file_path: str):
    entries = {}
    with open(plant_map_file_path, "r") as csv_file:
        reader = csv.reader(csv_file)
        next(reader)
        for row in reader:
            sddp_plant = SddpPlant()
            sddp_plant.system = row[0].strip().lower()
            sddp_plant.type = row[1].strip().lower()
            sddp_plant.name = row[2].strip()
            entry = PlantMapEntry()
            entry.plant = sddp_plant
            entry.weight = float(row[3])
            entry.elm_name = row[4].strip()

            if sddp_plant not in entries.keys():
                entries[sddp_plant] = [entry, ]
            else:
                entries[sddp_plant].append(entry)
    return entries


def _read_scenario_map(scenario_map_file_path: str):
    entries = {}
    with open(scenario_map_file_path, "r") as csv_file:
        reader = csv.reader(csv_file)
        next(reader)
        for row in reader:
            sddp_scenario = SddpScenario()
            sddp_scenario.stage = int(row[0])
            sddp_scenario.scenario = int(row[1])
            sddp_scenario.block = int(row[2])
            name = row[3].strip()
            entries[sddp_scenario] = name
    return entries


def _redistribute_weights(plant_map: dict):
    for plant in plant_map.keys():
        entries = plant_map[plant]
        total_weight = sum([entry.weight for entry in entries])
        for entry in entries:
            entry.weight /= total_weight


def _get_required_plant_types(plant_map: dict):
    types = set()
    for plant in plant_map.keys():
        types.add(plant.type)
    return types


def _load_graf_data(base_file_path: str, encoding: str):
    extensions_to_try = ".csv", ".hdr", ".dat"
    for ext in extensions_to_try:
        file_path = base_file_path + ext
        if os.path.exists(file_path):
            if _HAS_PANDAS:
                return psr.graf.load_as_dataframe(
                    file_path, encoding=encoding)
            else:
                ReaderClass = psr.graf.CsvReader if ext == ".csv" else \
                    psr.graf.BinReader
                obj = ReaderClass()
                obj.open(file_path, encoding=encoding)
                return obj
    return None


def _load_plant_types_generation(sddp_case_path: str, plant_types: set,
                                 encoding: str):
    generation_df = {}
    for plant_type in plant_types:
        base_file_name = os.path.join(sddp_case_path,
                                      _PLANT_TYPE_OUTPUT_MAP[plant_type])
        generation_df[plant_type] = _load_graf_data(base_file_name,
                                                    encoding=encoding)
    return generation_df


def _get_required_powerfactory_generators_names(plant_map: dict):
    generators = set()
    for plant in plant_map.keys():
        entries = plant_map[plant]
        for entry in entries:
            generators.add(entry.elm_name)
    return generators


def _get_powerfactory_objects_by_full_name(app, names: set):
    objects = {}
    for name in names:
        obj = app.GetCalcRelevantObjects(name)
        if obj is not None and len(obj) > 0:
            objects[name] = obj
    return objects


def main():
    # parse arguments
    # accepts an argument with the PF project name
    parser = argparse.ArgumentParser()
    parser.add_argument("-p", "--project", help="PowerFactory project name",
                        default="Transmission Example")
    parser.add_argument("-e", "--encoding", help="Result files encoding",
                        default="utf-8")
    # Sddp case path
    parser.add_argument("-sp", "--path", help="SDDP case path",
                        default=".")
    args = parser.parse_args()

    pf_project_name = args.project
    encoding = args.encoding
    sddp_case_path = args.path
    scenario_names_path = "scenarios_names.csv"
    if _DEBUG_PRINT:
        print("Reading scenarios_names")
    scenario_names = _read_scenario_map(scenario_names_path)

    if _DEBUG_PRINT:
        print("Reading durations")

    durations_df = _load_graf_data(os.path.join(sddp_case_path,
                                                _DURATION_FILE),
                                   encoding=encoding)

    if _DEBUG_PRINT:
        print("Reading plant -> elm map")
    plant_map_path = "plant_elm_map.csv"
    plant_map = _read_plant_map(plant_map_path)
    _redistribute_weights(plant_map)
    plant_types = _get_required_plant_types(plant_map)
    generation_df = _load_plant_types_generation(sddp_case_path, plant_types,
                                                 encoding=encoding)

    pf_generator_names = _get_required_powerfactory_generators_names(plant_map)

    if _DEBUG_PRINT:
        print("Starting powerfactory")
    app = powerfactory.GetApplication()
    app.ClearOutputWindow()
    app.PrintInfo("Starting")
    if _DEBUG_PRINT:
        print("Activating project", pf_project_name)
    app.ActivateProject(pf_project_name)

    if _DEBUG_PRINT:
        print("Loading PF generators")
    pf_generators = _get_powerfactory_objects_by_full_name(app,
                                                           pf_generator_names)
    if _DEBUG_PRINT:
        print("Loaded generators")
        print(pf_generators)
        print("Loaded Scenarios")
        print(scenario_names)
    for scenario, scenario_name in scenario_names.items():
        # Create scenario
        scenario_obj = create_scenario(app, scenario_name)
        if _DEBUG_PRINT:
            print("Creating scenario", scenario_name)
        app.PrintInfo("Creating scenario: " + scenario_name)
        scenario_obj[0].Activate()

        if _HAS_PANDAS:
            scn_tuple = scenario.stage, scenario.block
            scenario_duration_h = durations_df.loc[scenario.as_tuple(),
                                                   :][0][0]
        else:
            agents = durations_df.agents
            fixed_scenario = 1
            all_values = durations_df.read(scenario.stage, fixed_scenario,
                                           scenario.block)
            scenario_duration_h = all_values[0]
        units_conversion = 1000.0 / scenario_duration_h
        for elm_name, elm in pf_generators.items():
            if _DEBUG_PRINT:
                print("Setting generator", elm_name)
            elm_class = elm[0].GetClassName()
            attr_name = _ELM_CLASS_ATTRIBUTE_MAP[elm_class]
            attr_factor = _ELM_CLASS_ATTRIBUTE_FACTOR_MAP[elm_class]
            # search for related plant in plant map
            # TODO: maybe an inverse map is better
            plant = None
            weight = 1.0
            for sddp_plant, plant_map_entries in plant_map.items():
                for entry in plant_map_entries:
                    if elm_name == entry.elm_name:
                        plant = sddp_plant
                        weight = entry.weight
                        break
                if plant is not None:
                    break
            if plant is None:
                app.PrintInfo("Plant not found for generator: " + elm_name)
            else:
                if _DEBUG_PRINT:
                    print("Plant found:", plant.system, plant.type, plant.name)
                plant_type = plant.type
                plant_name = plant.name
                gen_type_df = generation_df[plant_type]
                if _HAS_PANDAS:
                    sddp_value = gen_type_df.loc[scenario.as_tuple(),
                                                 plant_name][0][0]
                else:
                    agents = gen_type_df.agents
                    all_values = gen_type_df.read(scenario.stage,
                                                  scenario.scenario,
                                                  scenario.block)
                    sddp_value = all_values[agents.index(plant_name)]
                value = sddp_value * units_conversion * weight * attr_factor
                if _DEBUG_PRINT:
                    print("Value read:", sddp_value, "Value assigned:", value)
                setattr(elm[0], attr_name, value)

        save_scenario = 1
        scenario_obj[0].Deactivate(save_scenario)
        if _DEBUG_PRINT:
            print("Scenario created", scenario_name)
        app.PrintInfo("Scenario created: " + scenario_name)
    print("Finished")
    app.PrintInfo("Finished")


if __name__ == "__main__":
    main()
