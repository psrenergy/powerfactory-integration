import argparse
import csv
import logging
import os
import sys
from dataclasses import dataclass

import psr.factory

logger = logging.getLogger("sddp_plant_dispatch")


_PLANT_TYPE_OUTPUT_MAP = {
    "hydro": "gerhid",
    "thermal": "gerter",
    "renewable": "gergnd",
    "battery": "gerbat",
    "csp": "cspgen",
    "injection": "powinj",
}

_DURATION_FILE = "duraci"

_SCENARIO_MAP_FILE = "scenario_names.csv"

_PLANT_MAP_FILE = "plant_elm_map.csv"

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
    "ElmPvsys": 0.001,  # TODO: confirm pgini unit for ElmPvsys
    "ElmXnet": 1.0,  # MW
}


class DispatchError(Exception):
    pass


@dataclass(frozen=True)
class SddpPlant:
    system: str
    type: str
    name: str


@dataclass(frozen=True)
class SddpScenario:
    stage: int
    scenario: int
    block: int

    def as_tuple(self):
        return self.stage, self.scenario, self.block


@dataclass
class PlantMapEntry:
    plant: SddpPlant
    weight: float
    elm_name: str


class _PowerFactoryLogHandler(logging.Handler):
    """Forwards log records to PowerFactory's output window."""
    def __init__(self, app):
        super().__init__(logging.INFO)
        self.app = app

    def emit(self, record):
        message = self.format(record)
        if record.levelno >= logging.ERROR:
            self.app.PrintError(message)
        elif record.levelno >= logging.WARNING:
            self.app.PrintWarn(message)
        else:
            self.app.PrintInfo(message)


def _open_csv(file_path: str):
    if not os.path.exists(file_path):
        raise DispatchError("Input file not found: " + file_path)
    return open(file_path, "r", encoding="utf-8-sig", newline="")


def _read_plant_map(plant_map_file_path: str):
    entries = {}
    with _open_csv(plant_map_file_path) as csv_file:
        reader = csv.reader(csv_file)
        next(reader)
        for line, row in enumerate(reader, start=2):
            if not row:
                continue
            plant_type = row[1].strip().lower()
            if plant_type not in _PLANT_TYPE_OUTPUT_MAP:
                raise DispatchError(
                    "{}:{}: unknown plant type '{}'. Supported types: {}"
                    .format(plant_map_file_path, line, plant_type,
                            ", ".join(_PLANT_TYPE_OUTPUT_MAP)))
            sddp_plant = SddpPlant(system=row[0].strip().lower(),
                                   type=plant_type,
                                   name=row[2].strip())
            entry = PlantMapEntry(plant=sddp_plant, weight=float(row[3]),
                                  elm_name=row[4].strip())
            entries.setdefault(sddp_plant, []).append(entry)
    return entries


def _read_scenario_map(scenario_map_file_path: str):
    entries = {}
    with _open_csv(scenario_map_file_path) as csv_file:
        reader = csv.reader(csv_file)
        next(reader)
        for row in reader:
            if not row:
                continue
            sddp_scenario = SddpScenario(stage=int(row[0]),
                                         scenario=int(row[1]),
                                         block=int(row[2]))
            entries[sddp_scenario] = row[3].strip()
    return entries


def _redistribute_weights(plant_map: dict):
    for plant, entries in plant_map.items():
        total_weight = sum(entry.weight for entry in entries)
        if total_weight <= 0.0:
            raise DispatchError(
                "Total weight of plant '{}' ({}, system '{}') must be "
                "positive.".format(plant.name, plant.type, plant.system))
        for entry in entries:
            entry.weight /= total_weight


def _build_elm_map(plant_map: dict):
    """Maps PowerFactory element names to the plant entries feeding them."""
    elm_map = {}
    for entries in plant_map.values():
        for entry in entries:
            elm_map.setdefault(entry.elm_name, []).append(entry)
    return elm_map


class SddpResult:
    """SDDP result file values indexed by (stage, scenario, block)."""
    def __init__(self, df, file_path: str):
        self.file_path = file_path
        index_names = df.index_names
        block_column = "block" if "block" in index_names else "hour"
        stages = df["stage"]
        blocks = df[block_column]
        # Files with a single scenario have no scenario index column
        self.has_scenarios = "scenario" in index_names
        scenarios = df["scenario"] if self.has_scenarios else [1] * len(df)
        self.rows = {}
        for row, key in enumerate(zip(stages, scenarios, blocks)):
            self.rows[tuple(int(k) for k in key)] = row
        self.agents = list(df.column_names)
        self.columns = {agent: df[agent] for agent in self.agents}

    def value(self, scenario, agent=None):
        stage, scenario_number, block = scenario.as_tuple()
        if not self.has_scenarios:
            scenario_number = 1
        column = self.agents[0] if agent is None else agent
        if column not in self.columns:
            raise DispatchError("Agent '{}' not found in {}".format(
                column, self.file_path))
        key = (stage, scenario_number, block)
        if key not in self.rows:
            raise DispatchError(
                "Stage {}, scenario {}, block {} not found in {}".format(
                    *key, self.file_path))
        return float(self.columns[column][self.rows[key]])


def _load_sddp_result(base_file_path: str):
    extensions_to_try = ".csv", ".hdr"
    for ext in extensions_to_try:
        file_path = base_file_path + ext
        if os.path.exists(file_path):
            df = psr.factory.load_dataframe(
                file_path, options={"IndexFormat": "stage"})
            return SddpResult(df, file_path)
    raise DispatchError("SDDP result file not found: " + base_file_path)


def _load_plant_types_generation(sddp_case_path: str, plant_types: set):
    generation_df = {}
    for plant_type in plant_types:
        base_file_name = os.path.join(sddp_case_path,
                                      _PLANT_TYPE_OUTPUT_MAP[plant_type])
        generation_df[plant_type] = _load_sddp_result(base_file_name)
    return generation_df


def _get_powerfactory_elements(app, names):
    elements = {}
    for name in sorted(names):
        objs = app.GetCalcRelevantObjects(name)
        if not objs:
            logger.warning("PowerFactory element not found: %s", name)
            continue
        if len(objs) > 1:
            logger.warning("Multiple PowerFactory elements match %s; "
                           "all of them will be set.", name)
        supported = []
        for obj in objs:
            elm_class = obj.GetClassName()
            if elm_class in _ELM_CLASS_ATTRIBUTE_MAP:
                supported.append(obj)
            else:
                logger.warning("Unsupported element class %s for %s; "
                               "skipping.", elm_class, name)
        if supported:
            elements[name] = supported
    return elements


def _get_or_create_scenario(app, scenario_name: str):
    scen_folder = app.GetProjectFolder("scen")
    existing = scen_folder.GetContents(scenario_name + ".IntScenario")
    if existing:
        logger.info("Updating existing scenario: %s", scenario_name)
        return existing[0]
    logger.info("Creating scenario: %s", scenario_name)
    return scen_folder.CreateObject("IntScenario", scenario_name)


def _compute_elements_power(scenario, scenario_name, durations,
                            generation_df, elm_map, elm_names):
    """Returns the power in MW of each element for a SDDP scenario."""
    duration_h = durations.value(scenario)
    if duration_h <= 0.0:
        raise DispatchError("Non-positive block duration for scenario "
                            + scenario_name)
    # SDDP generation in GWh per block -> MW
    units_conversion = 1000.0 / duration_h

    elements_power = {}
    for elm_name in elm_names:
        power_mw = 0.0
        for entry in elm_map[elm_name]:
            plant = entry.plant
            sddp_value = generation_df[plant.type].value(scenario, plant.name)
            power_mw += sddp_value * units_conversion * entry.weight
        elements_power[elm_name] = power_mw
    return elements_power


def _print_dry_run(scenario_names, durations, generation_df, elm_map):
    writer = csv.writer(sys.stdout, lineterminator="\n")
    writer.writerow(["Scenario Name", "Element", "Power (MW)"])
    for scenario, scenario_name in scenario_names.items():
        elements_power = _compute_elements_power(
            scenario, scenario_name, durations, generation_df, elm_map,
            sorted(elm_map.keys()))
        for elm_name, power_mw in elements_power.items():
            writer.writerow([scenario_name, elm_name,
                             "{:.6f}".format(power_mw)])


def _apply_scenario(app, scenario, scenario_name, durations, generation_df,
                    elm_map, pf_elements):
    scenario_obj = _get_or_create_scenario(app, scenario_name)
    scenario_obj.Activate()

    elements_power = _compute_elements_power(
        scenario, scenario_name, durations, generation_df, elm_map,
        pf_elements.keys())
    for elm_name, objs in pf_elements.items():
        power_mw = elements_power[elm_name]
        for obj in objs:
            elm_class = obj.GetClassName()
            value = power_mw * _ELM_CLASS_ATTRIBUTE_FACTOR_MAP[elm_class]
            logger.debug("%s: %s = %f", elm_name,
                         _ELM_CLASS_ATTRIBUTE_MAP[elm_class], value)
            setattr(obj, _ELM_CLASS_ATTRIBUTE_MAP[elm_class], value)

    save_scenario = 1
    scenario_obj.Deactivate(save_scenario)
    logger.info("Scenario saved: %s", scenario_name)


def _parse_args():
    parser = argparse.ArgumentParser(
        description="Distribute SDDP plant generation in PowerFactory "
                    "generators, one operation scenario per SDDP "
                    "stage/scenario/block.")
    parser.add_argument("-p", "--project", help="PowerFactory project name",
                        default="Transmission Example")
    parser.add_argument("-sp", "--path", help="SDDP case path", default=".")
    parser.add_argument("--scenarios",
                        help="Scenario names CSV file (default: {} in the "
                             "SDDP case path)".format(_SCENARIO_MAP_FILE))
    parser.add_argument("--plant-map",
                        help="Plant/element map CSV file (default: {} in the "
                             "SDDP case path)".format(_PLANT_MAP_FILE))
    parser.add_argument("-n", "--dry-run", action="store_true",
                        help="Print the power of each element per scenario "
                             "as CSV without connecting to PowerFactory")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Print debug messages")
    return parser.parse_args()


def main():
    args = _parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s: %(message)s")

    sddp_case_path = args.path
    scenario_map_path = args.scenarios or os.path.join(sddp_case_path,
                                                       _SCENARIO_MAP_FILE)
    plant_map_path = args.plant_map or os.path.join(sddp_case_path,
                                                    _PLANT_MAP_FILE)

    logger.debug("Reading %s", scenario_map_path)
    scenario_names = _read_scenario_map(scenario_map_path)
    logger.debug("Reading %s", plant_map_path)
    plant_map = _read_plant_map(plant_map_path)
    _redistribute_weights(plant_map)
    elm_map = _build_elm_map(plant_map)

    durations = _load_sddp_result(os.path.join(sddp_case_path,
                                               _DURATION_FILE))
    plant_types = {plant.type for plant in plant_map}
    generation_df = _load_plant_types_generation(sddp_case_path, plant_types)

    if args.dry_run:
        _print_dry_run(scenario_names, durations, generation_df, elm_map)
        return

    # Imported here so dry runs work without PowerFactory
    import powerfactory
    app = powerfactory.GetApplication()
    if app is None:
        raise DispatchError("Could not connect to PowerFactory.")
    app.ClearOutputWindow()
    pf_handler = _PowerFactoryLogHandler(app)
    pf_handler.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(pf_handler)

    logger.info("Activating project: %s", args.project)
    if app.ActivateProject(args.project) != 0:
        raise DispatchError("Could not activate PowerFactory project: "
                            + args.project)

    pf_elements = _get_powerfactory_elements(app, elm_map.keys())
    if not pf_elements:
        raise DispatchError("None of the mapped PowerFactory elements "
                            "were found.")

    for scenario, scenario_name in scenario_names.items():
        _apply_scenario(app, scenario, scenario_name, durations,
                        generation_df, elm_map, pf_elements)
    logger.info("Finished")


if __name__ == "__main__":
    try:
        main()
    except DispatchError as e:
        logger.error(str(e))
        sys.exit(1)
