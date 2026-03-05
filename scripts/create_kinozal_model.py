"""
Создание модели OpenStudio для большого кинозала (135 мест)
СПб ГБУК "Петербург-кино", Ярославский пр., д. 55, лит. А

Источник: Проектная документация 2017/23-ИОС4-О, ООО «АБИ 1», 2017 г.

Параметры зала:
  Площадь: 133 м², Высота: 4.85 м, Объём: 645 м³, 135 мест
  П1 (приток): 4050 м³/ч, В1 (вытяжка): 4050 м³/ч, кратность: 6.28 ч⁻¹
  Вентустановка: AIRNED-M6L, вентилятор: 5425 м³/ч / 500 Па / 1.97 кВт
  Теплоноситель: 80-60°C, рекуператор роторный RRS

Шесть сценариев:
  Base     — фиксированный расход 4050 м³/ч, рекуператор η=80%
  DCV      — управление по CO₂ (уставка 700 ppm), рекуператор η=80%
  Base_HP  — фиксированный расход, тепловой насос, рекуператор η=80%
  DCV_HP   — DCV (уставка 700 ppm) + тепловой насос, рекуператор η=80%
  Base_Rec — фиксированный расход, высокоэффективный рекуператор η=90%
  DCV_Rec  — DCV (уставка 700 ppm), высокоэффективный рекуператор η=90%

Изменения DCV относительно первоначальной версии:
  - Уставка CO₂: 800 → 700 ppm (активация до пересечения порога «Отличное»)
  - Минимальный расход НВ: 0 → 0.5625 м³/с (2025 м³/ч, 50%)
  - Минимальная доля терминала: 20% → 50%
  - Минимальная доля вентилятора: 20% → 50%
  Это обеспечивает категорию «Отличное» (CO₂ < 800 ppm) при той же энергоэкономии.
"""

import sys, os, json
sys.path.insert(0, "/Applications/OpenStudio-3.10.0/Python")
import openstudio

ROOT_DIR = "/Users/m/Documents/Openstudio"
WEATHER_PATH = os.path.join(ROOT_DIR, "weather", "Saint-Petersburg.epw")
MODELS_DIR = os.path.join(ROOT_DIR, "models")

# ─────────────────────────────────────────────────────────────────────────────
def make_const_schedule(model, name, value):
    s = openstudio.model.ScheduleConstant(model)
    s.setName(name)
    s.setValue(value)
    return s

def make_ruleset_schedule(model, name, weekday_vals, weekend_vals=None,
                           lower=0.0, upper=1.0, unit="Dimensionless"):
    """24-значный массив (индекс=час 0..23, значение действует ДО конца этого часа)."""
    stl = openstudio.model.ScheduleTypeLimits(model)
    stl.setName(f"{name}_Limits")
    stl.setLowerLimitValue(lower)
    stl.setUpperLimitValue(upper)
    stl.setNumericType("Continuous")
    stl.setUnitType(unit)

    sch = openstudio.model.ScheduleRuleset(model)
    sch.setName(name)
    sch.setScheduleTypeLimits(stl)

    def fill_day(day_sch, vals):
        for h, v in enumerate(vals):
            day_sch.addValue(openstudio.Time(0, h + 1, 0, 0), float(v))

    fill_day(sch.defaultDaySchedule(), weekday_vals)

    if weekend_vals:
        rule = openstudio.model.ScheduleRule(sch)
        rule.setApplySaturday(True)
        rule.setApplySunday(True)
        fill_day(rule.daySchedule(), weekend_vals)

    return sch

def add_mat(model, name, thick, cond, dens, cp, solar_abs=0.7):
    m = openstudio.model.StandardOpaqueMaterial(model)
    m.setName(name)
    m.setThickness(thick)
    m.setConductivity(cond)
    m.setDensity(dens)
    m.setSpecificHeat(cp)
    m.setSolarAbsorptance(solar_abs)
    m.setThermalAbsorptance(0.9)
    return m

def make_construction(model, name, layers):
    c = openstudio.model.Construction(model)
    c.setName(name)
    for i, mat in enumerate(layers):
        c.insertLayer(i, mat)
    return c

def add_surface(model, space, name, stype, verts, construction=None, obc="Outdoors"):
    pts = openstudio.Point3dVector()
    for v in verts:
        pts.append(openstudio.Point3d(*v))
    s = openstudio.model.Surface(pts, model)
    s.setName(name)
    s.setSpace(space)
    s.setSurfaceType(stype)
    s.setOutsideBoundaryCondition(obc)
    if construction:
        s.setConstruction(construction)
    return s

# ─────────────────────────────────────────────────────────────────────────────
def build_model(dcv=False, heat_pump=False, recuperator_high=False):
    if heat_pump:
        tag = "DCV_HP" if dcv else "Base_HP"
    elif recuperator_high:
        tag = "DCV_Rec" if dcv else "Base_Rec"
    else:
        tag = "DCV" if dcv else "Base"
    model = openstudio.model.Model()

    # ── 1. РАСПОЛОЖЕНИЕ: Санкт-Петербург ──────────────────────────────────────
    site = model.getSite()
    site.setName("SPb_Pulkovo")
    site.setLatitude(59.80)
    site.setLongitude(30.26)
    site.setTimeZone(3.0)
    site.setElevation(16.0)

    if os.path.exists(WEATHER_PATH):
        try:
            ef = openstudio.EpwFile(openstudio.path(WEATHER_PATH))
            openstudio.model.WeatherFile.setWeatherFile(model, ef)
            print(f"  [OK] EPW: {WEATHER_PATH}")
        except Exception as e:
            print(f"  [!] EPW error: {e}")
    else:
        print(f"  [!] EPW not found: {WEATHER_PATH}")

    # ── 2. ПАРАМЕТРЫ СИМУЛЯЦИИ ────────────────────────────────────────────────
    sc = model.getSimulationControl()
    sc.setRunSimulationforWeatherFileRunPeriods(True)
    sc.setRunSimulationforSizingPeriods(True)
    sc.setDoZoneSizingCalculation(True)
    sc.setDoSystemSizingCalculation(True)
    sc.setDoPlantSizingCalculation(True)

    rp = model.getRunPeriod()
    rp.setBeginMonth(1);  rp.setBeginDayOfMonth(1)
    rp.setEndMonth(12);   rp.setEndDayOfMonth(31)

    # Отслеживание CO₂ (обязательно нужно задать расписание наружной концентрации)
    zacb = model.getZoneAirContaminantBalance()
    zacb.setCarbonDioxideConcentration(True)
    outdoor_co2_sch = make_const_schedule(model, "Наружный_CO2_400ppm", 400.0)
    zacb.setOutdoorCarbonDioxideSchedule(outdoor_co2_sch)

    # ── 3. МАТЕРИАЛЫ И КОНСТРУКЦИИ ────────────────────────────────────────────
    # Стены: штукатурка + бетон 500 + минвата 100 + штукатурка
    plaster   = add_mat(model, "Штукатурка_20мм",    0.020, 0.70, 1600, 840)
    concrete  = add_mat(model, "Бетон_500мм",        0.500, 1.74, 2200, 840)
    mw100     = add_mat(model, "МинВата_100мм",      0.100, 0.042,  80, 840)
    mw150     = add_mat(model, "МинВата_150мм",      0.150, 0.042,  80, 840)
    slab220   = add_mat(model, "ЖБ_плита_220мм",     0.220, 1.74, 2500, 840)
    screed100 = add_mat(model, "Стяжка_100мм",       0.100, 0.93, 1800, 840)

    c_wall_ext = make_construction(model, "Стена_наружная",
                                   [plaster, concrete, mw100, plaster])
    c_wall_int = make_construction(model, "Стена_внутренняя",
                                   [plaster, concrete, plaster])
    c_roof     = make_construction(model, "Покрытие_зала",
                                   [mw150, slab220])
    c_floor    = make_construction(model, "Пол_зала",
                                   [screed100, slab220])

    # ── 4. ТЕРМИЧЕСКАЯ ЗОНА ───────────────────────────────────────────────────
    tz = openstudio.model.ThermalZone(model)
    tz.setName(f"Кинозал_135мест_{tag}")

    space = openstudio.model.Space(model)
    space.setName(f"Зал_{tag}")
    space.setThermalZone(tz)

    # ── 5. ГЕОМЕТРИЯ ─────────────────────────────────────────────────────────
    # Зал 133 м²: 15.68 × 8.48 × 4.85 м
    L, W, H = 15.68, 8.48, 4.85

    # Пол (ЭП: вершины по ЧС при взгляде сверху = нормаль вниз)
    fl = add_surface(model, space, "Пол", "Floor",
                     [(0,0,0),(0,W,0),(L,W,0),(L,0,0)], c_floor, "Ground")

    # Потолок (ЭП: вершины против ЧС при взгляде сверху = нормаль вверх)
    rf = add_surface(model, space, "Потолок", "RoofCeiling",
                     [(0,0,H),(L,0,H),(L,W,H),(0,W,H)], c_roof)
    rf.setSunExposure("NoSun"); rf.setWindExposure("NoWind")

    # Стены (наружные). Порядок вершин задан так, чтобы нормали были направлены наружу.
    add_surface(model, space, "Стена_Ю", "Wall",
                [(L,0,0),(L,0,H),(0,0,H),(0,0,0)], c_wall_ext)
    add_surface(model, space, "Стена_С", "Wall",
                [(0,W,0),(0,W,H),(L,W,H),(L,W,0)], c_wall_ext)
    add_surface(model, space, "Стена_З", "Wall",
                [(0,0,0),(0,0,H),(0,W,H),(0,W,0)], c_wall_int)
    add_surface(model, space, "Стена_В", "Wall",
                [(L,W,0),(L,W,H),(L,0,H),(L,0,0)], c_wall_int)

    print(f"  [OK] Геометрия: {L}×{W}×{H}м = {L*W:.1f}м²")

    # ── 6. РАСПИСАНИЯ ─────────────────────────────────────────────────────────
    # Кинотеатр: работа с 9:00 до 24:00, сеансы ~2ч каждые 2-3ч
    # Заполняемость по часам (0=пусто, 1=полный зал 135 чел)
    occ_wd = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,
              0.1, 0.1, 0.2,
              0.85, 0.85,   # 12-13: 1-й сеанс
              0.85, 0.85,   # 14-15: 2-й сеанс
              0.15,
              0.80, 0.80,   # 17-18: 3-й сеанс
              0.15,
              0.92, 0.92,   # 20-21: 4-й сеанс (пик)
              0.20,
              0.82, 0.82]   # 23-24: поздний

    occ_we = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,
              0.1, 0.1,
              0.88, 0.88,   # 11-12: утренний
              0.92, 0.92,   # 13-14: дневной
              0.20,
              0.90, 0.90,   # 16-17: дневной
              0.20,
              0.95, 0.95,   # 19-20: вечерний
              0.30,
              0.80, 0.80,   # 22-23: поздний
              0.0]

    sch_occ  = make_ruleset_schedule(model, f"Заполняемость_{tag}", occ_wd, occ_we)
    sch_light = make_ruleset_schedule(model, f"Освещение_{tag}",
                                      [min(1.0, v+0.05) if v>0 else 0.0 for v in occ_wd],
                                      [min(1.0, v+0.05) if v>0 else 0.0 for v in occ_we])
    sch_vent = make_ruleset_schedule(model, f"Вентиляция_{tag}",
                                     [0.0]*8 + [1.0]*16,
                                     [0.0]*8 + [1.0]*16)
    sch_act  = make_const_schedule(model, f"Активность_{tag}", 115.0)  # Вт/чел (сидя)
    sch_htg  = make_const_schedule(model, f"Уставка_отопл_{tag}", 18.0)
    sch_clg  = make_const_schedule(model, f"Уставка_охл_{tag}",   26.0)
    sch_supply_t = make_const_schedule(model, f"t_приточного_{tag}", 18.0)

    # ── 7. НАГРУЗКИ ───────────────────────────────────────────────────────────
    # Люди: 135 зрителей
    pd = openstudio.model.PeopleDefinition(model)
    pd.setName(f"Зрители_{tag}")
    pd.setNumberofPeople(135)
    pd.setFractionRadiant(0.30)
    pd.setSensibleHeatFraction(0.60)
    pd.setCarbonDioxideGenerationRate(0.0000049)   # 4.9e-6 м³/с·чел (ASHRAE)

    ppl = openstudio.model.People(pd)
    ppl.setName(f"Люди_{tag}")
    ppl.setSpace(space)
    ppl.setNumberofPeopleSchedule(sch_occ)
    ppl.setActivityLevelSchedule(sch_act)

    # Освещение: 8 Вт/м² (рабочее + аварийное)
    ld = openstudio.model.LightsDefinition(model)
    ld.setName(f"Освещение_дефиниция_{tag}")
    ld.setWattsperSpaceFloorArea(8.0)

    lgt = openstudio.model.Lights(ld)
    lgt.setName(f"Освещение_{tag}")
    lgt.setSpace(space)
    lgt.setSchedule(sch_light)

    # Кинопроектор: 10 кВт
    ed = openstudio.model.ElectricEquipmentDefinition(model)
    ed.setName(f"Проектор_дефиниция_{tag}")
    ed.setDesignLevel(10000.0)
    ed.setFractionRadiant(0.20)
    ed.setFractionLatent(0.0)

    eq = openstudio.model.ElectricEquipment(ed)
    eq.setName(f"Проектор_{tag}")
    eq.setSpace(space)
    eq.setSchedule(sch_light)

    # Инфильтрация (задана минимально - здание плотное, сцена кинозала)
    inf = openstudio.model.SpaceInfiltrationDesignFlowRate(model)
    inf.setName(f"Инфильтрация_{tag}")
    inf.setFlowperExteriorSurfaceArea(0.0001)   # ~0.36 м³/ч·м² — плотный фасад
    inf.setSpace(space)

    # ── 8. ТРЕБОВАНИЯ К НАРУЖНОМУ ВОЗДУХУ ────────────────────────────────────
    # Проектный расход: 4050 м³/ч = 30 м³/ч·чел = 8.33 л/с·чел
    dsoa = openstudio.model.DesignSpecificationOutdoorAir(model)
    dsoa.setName(f"ОА_спец_{tag}")

    if dcv:
        # DCV с нижним пределом НВ:
        # max(0.619 м³/с, 8.33 л/с·чел * N), где 0.619 м³/с = 55% от номинала.
        dsoa.setOutdoorAirMethod("Maximum")
        dsoa.setOutdoorAirFlowRate(0.619)
        dsoa.setOutdoorAirFlowperPerson(0.00833)    # 8.33 л/с/чел = 30 м³/ч/чел
        dsoa.setOutdoorAirFlowperFloorArea(0.0)
    else:
        # Фиксированный расход (независимо от заполненности)
        dsoa.setOutdoorAirMethod("Sum")
        dsoa.setOutdoorAirFlowRate(1.125)           # 4050/3600 = 1.125 м³/с
        dsoa.setOutdoorAirFlowperPerson(0.0)
        dsoa.setOutdoorAirFlowperFloorArea(0.0)

    space.setDesignSpecificationOutdoorAir(dsoa)

    # ── 9. ТЕРМОСТАТ ──────────────────────────────────────────────────────────
    tstat = openstudio.model.ThermostatSetpointDualSetpoint(model)
    tstat.setName(f"Термостат_{tag}")
    tstat.setHeatingSetpointTemperatureSchedule(sch_htg)
    tstat.setCoolingSetpointTemperatureSchedule(sch_clg)
    tz.setThermostatSetpointDualSetpoint(tstat)

    # ── 10. СИСТЕМА ВЕНТИЛЯЦИИ П1В1 (AIRNED-M6L) ─────────────────────────────
    air_loop = openstudio.model.AirLoopHVAC(model)
    air_loop.setName(f"П1В1_AIRNED_M6L_{tag}")
    air_loop.setDesignSupplyAirFlowRate(1.125)   # 4050/3600 м³/с

    sup_node = air_loop.supplyOutletNode()

    # -- Водяной калорифер (теплоноситель 80-60°C) ----------------------------
    htg_coil = openstudio.model.CoilHeatingWater(model)
    htg_coil.setName(f"Калорифер_П1_{tag}")
    htg_coil.addToNode(sup_node)

    # -- Приточный вентилятор -------------------------------------------------
    # Проектные данные: 5425 м³/ч, 500 Па, η_tot≈65%, N=1.97 кВт
    if dcv:
        fan = openstudio.model.FanVariableVolume(model)
        fan.setName(f"Вент_П1_VAV_{tag}")
        fan.setMaximumFlowRate(1.125)        # 4050/3600 м³/с (проект)
        fan.setFanTotalEfficiency(0.65)
        fan.setPressureRise(500.0)
        fan.setMotorEfficiency(0.90)
        fan.setMotorInAirstreamFraction(1.0)
        fan.setFanPowerMinimumFlowRateInputMethod("Fraction")
        fan.setFanPowerMinimumFlowFraction(0.55)  # мин. 55% (2228 м³/ч) → CO₂ «Отличное»
    else:
        fan = openstudio.model.FanConstantVolume(model)
        fan.setName(f"Вент_П1_CV_{tag}")
        fan.setMaximumFlowRate(1.125)
        fan.setFanTotalEfficiency(0.65)
        fan.setPressureRise(500.0)
        fan.setMotorEfficiency(0.90)
        fan.setMotorInAirstreamFraction(1.0)
    fan.setAvailabilitySchedule(sch_vent)
    fan.addToNode(sup_node)

    # -- Уставка температуры подачи -------------------------------------------
    sp_mgr = openstudio.model.SetpointManagerScheduled(model, "Temperature", sch_supply_t)
    sp_mgr.setName(f"СПМ_t_подачи_{tag}")
    sp_mgr.addToNode(sup_node)

    # -- Система наружного воздуха --------------------------------------------
    oa_ctrl = openstudio.model.ControllerOutdoorAir(model)
    oa_ctrl.setName(f"Контроллер_НВ_{tag}")
    # Для DCV не задаём жёсткий минимум OA на контроллере, чтобы не подавлять
    # запрос Controller:MechanicalVentilation и исключить конфликт Min OA fraction.
    oa_ctrl.setMinimumOutdoorAirFlowRate(0.0)
    oa_ctrl.setMaximumOutdoorAirFlowRate(1.125)
    oa_ctrl.setEconomizerControlType("NoEconomizer")

    if dcv:
        # Включаем DCV через ControllerMechanicalVentilation
        mech_vent = oa_ctrl.controllerMechanicalVentilation()
        mech_vent.setDemandControlledVentilation(True)
        mech_vent.setSystemOutdoorAirMethod("ZoneSum")

    oa_sys = openstudio.model.AirLoopHVACOutdoorAirSystem(model, oa_ctrl)
    oa_sys.setName(f"ОА_система_{tag}")
    oa_sys.addToNode(sup_node)

    # -- Роторный рекуператор на стороне наружного воздуха --------------------
    # Размещается на outboardOANode (до смешивания с рециркуляцией),
    # secondary-сторона подключается к вытяжному потоку автоматически.
    hrv = openstudio.model.HeatExchangerAirToAirSensibleAndLatent(model)
    if recuperator_high:
        # Высокоэффективный электрический роторный рекуператор η_sens=90%, η_lat=82%
        hrv.setName(f"Рекуператор_HiEff_{tag}")
        hrv.setHeatExchangerType("Rotary")
        hrv.setSensibleEffectivenessat100HeatingAirFlow(0.90)
        hrv.setSensibleEffectivenessat75HeatingAirFlow(0.93)
        hrv.setLatentEffectivenessat100HeatingAirFlow(0.82)
        hrv.setLatentEffectivenessat75HeatingAirFlow(0.85)
        hrv.setSensibleEffectivenessat100CoolingAirFlow(0.85)
        hrv.setSensibleEffectivenessat75CoolingAirFlow(0.88)
        hrv.setLatentEffectivenessat100CoolingAirFlow(0.78)
        hrv.setLatentEffectivenessat75CoolingAirFlow(0.81)
    else:
        # Стандартный роторный рекуператор RRS η_sens=80%, η_lat=72%
        hrv.setName(f"Рекуператор_RRS_{tag}")
        hrv.setHeatExchangerType("Rotary")
        hrv.setSensibleEffectivenessat100HeatingAirFlow(0.80)
        hrv.setSensibleEffectivenessat75HeatingAirFlow(0.83)
        hrv.setLatentEffectivenessat100HeatingAirFlow(0.72)
        hrv.setLatentEffectivenessat75HeatingAirFlow(0.75)
        hrv.setSensibleEffectivenessat100CoolingAirFlow(0.75)
        hrv.setSensibleEffectivenessat75CoolingAirFlow(0.78)
        hrv.setLatentEffectivenessat100CoolingAirFlow(0.68)
        hrv.setLatentEffectivenessat75CoolingAirFlow(0.71)
    hrv.setEconomizerLockout(True)
    oa_outboard = oa_sys.outboardOANode()
    if oa_outboard.is_initialized():
        hrv.addToNode(oa_outboard.get())
        print(f"  [OK] Рекуператор добавлен в OA-систему: {hrv.name()}")

    # -- Терминал зоны --------------------------------------------------------
    if dcv:
        term = openstudio.model.AirTerminalSingleDuctVAVNoReheat(model, sch_vent)
        term.setName(f"Терминал_VAV_{tag}")
        term.setMaximumAirFlowRate(1.125)
        term.setZoneMinimumAirFlowInputMethod("Constant")
        term.setConstantMinimumAirFlowFraction(0.55)   # 55% мин (2228 м³/ч) → CO₂ «Отличное»
    else:
        term = openstudio.model.AirTerminalSingleDuctConstantVolumeNoReheat(model, sch_vent)
        term.setName(f"Терминал_CAV_{tag}")
        term.setMaximumAirFlowRate(1.125)

    air_loop.addBranchForZone(tz, term)

    # ── 11. CO₂ КОНТРОЛЛЕР ЗОНЫ (только DCV) ─────────────────────────────────
    if dcv:
        co2_sp_sch = make_const_schedule(model, "CO2_уставка_700ppm", 700.0)
        co2_avail_sch = make_ruleset_schedule(model, "CO2_ctrl_доступность",
                                              [0.0]*8 + [1.0]*16,
                                              [0.0]*8 + [1.0]*16)
        co2_ctrl = openstudio.model.ZoneControlContaminantController(model)
        co2_ctrl.setName(f"CO2_контроллер_{tag}")
        co2_ctrl.setCarbonDioxideControlAvailabilitySchedule(co2_avail_sch)
        co2_ctrl.setCarbonDioxideSetpointSchedule(co2_sp_sch)
        # Привязываем к зоне
        tz.setZoneControlContaminantController(co2_ctrl)

    # ── 12. КОНТУР ТЕПЛОСНАБЖЕНИЯ ────────────────────────────────────────────
    pl = openstudio.model.PlantLoop(model)
    pl.setName(f"ИТП_{tag}")
    if heat_pump:
        pl.setMaximumLoopTemperature(55.0)
        pl.setMinimumLoopTemperature(30.0)
        load_loop_setpoint = 45.0
    else:
        pl.setMaximumLoopTemperature(80.0)
        pl.setMinimumLoopTemperature(60.0)
        load_loop_setpoint = 80.0

    pl_sp = make_const_schedule(model, f"Уставка_ИТП_{tag}", load_loop_setpoint)
    pl_mgr = openstudio.model.SetpointManagerScheduled(model, "Temperature", pl_sp)
    pl_mgr.addToNode(pl.supplyOutletNode())

    pump = openstudio.model.PumpConstantSpeed(model)
    pump.setName(f"Насос_ИТП_{tag}")
    pump.setRatedPumpHead(30000.0)
    pump.setMotorEfficiency(0.85)
    pump.addToNode(pl.supplyInletNode())

    if heat_pump:
        # Нагрев через тепловой насос: load-side подключен к контуру калорифера,
        # source-side к отдельному источниковому контуру с фиксированной температурой.
        hp_cap_ft = openstudio.model.CurveBiquadratic(model)
        hp_cap_ft.setName(f"ТН_{tag}_CapModFT")
        hp_cap_ft.setCoefficient1Constant(1.0)
        hp_cap_ft.setCoefficient2x(0.0)
        hp_cap_ft.setCoefficient3xPOW2(0.0)
        hp_cap_ft.setCoefficient4y(0.0)
        hp_cap_ft.setCoefficient5yPOW2(0.0)
        hp_cap_ft.setCoefficient6xTIMESY(0.0)
        hp_cap_ft.setMinimumValueofx(-100.0)
        hp_cap_ft.setMaximumValueofx(100.0)
        hp_cap_ft.setMinimumValueofy(-100.0)
        hp_cap_ft.setMaximumValueofy(100.0)
        hp_cap_ft.setInputUnitTypeforX("Temperature")
        hp_cap_ft.setInputUnitTypeforY("Temperature")
        hp_cap_ft.setOutputUnitType("Dimensionless")

        hp_eir_ft = openstudio.model.CurveBiquadratic(model)
        hp_eir_ft.setName(f"ТН_{tag}_EIRModFT")
        hp_eir_ft.setCoefficient1Constant(1.0)
        hp_eir_ft.setCoefficient2x(0.0)
        hp_eir_ft.setCoefficient3xPOW2(0.0)
        hp_eir_ft.setCoefficient4y(0.0)
        hp_eir_ft.setCoefficient5yPOW2(0.0)
        hp_eir_ft.setCoefficient6xTIMESY(0.0)
        hp_eir_ft.setMinimumValueofx(-100.0)
        hp_eir_ft.setMaximumValueofx(100.0)
        hp_eir_ft.setMinimumValueofy(-100.0)
        hp_eir_ft.setMaximumValueofy(100.0)
        hp_eir_ft.setInputUnitTypeforX("Temperature")
        hp_eir_ft.setInputUnitTypeforY("Temperature")
        hp_eir_ft.setOutputUnitType("Dimensionless")

        hp_eir_fplr = openstudio.model.CurveQuadratic(model)
        hp_eir_fplr.setName(f"ТН_{tag}_EIRModFPLR")
        hp_eir_fplr.setCoefficient1Constant(1.0)
        hp_eir_fplr.setCoefficient2x(0.0)
        hp_eir_fplr.setCoefficient3xPOW2(0.0)
        hp_eir_fplr.setMinimumValueofx(0.0)
        hp_eir_fplr.setMaximumValueofx(1.0)
        hp_eir_fplr.setInputUnitTypeforX("Dimensionless")
        hp_eir_fplr.setOutputUnitType("Dimensionless")

        hp = openstudio.model.HeatPumpPlantLoopEIRHeating(model)
        hp.setName(f"ТН_источник_{tag}")
        hp.setReferenceCoefficientofPerformance(2.5)
        hp.setSizingFactor(1.0)
        hp.setControlType("Setpoint")
        hp.setFlowMode("ConstantFlow")
        hp.setMinimumPartLoadRatio(0.0)
        hp.setHeatPumpSizingMethod("HeatingCapacity")
        hp.setHeatingToCoolingCapacitySizingRatio(1.0)
        hp.setMinimumSourceInletTemperature(-100.0)
        hp.setMaximumSourceInletTemperature(-50.0)
        hp.setCapacityModifierFunctionofTemperatureCurve(hp_cap_ft)
        hp.setElectricInputtoOutputRatioModifierFunctionofTemperatureCurve(hp_eir_ft)
        hp.setElectricInputtoOutputRatioModifierFunctionofPartLoadRatioCurve(hp_eir_fplr)

        src = openstudio.model.PlantLoop(model)
        src.setName(f"Источник_ТН_{tag}")
        src.setMaximumLoopTemperature(25.0)
        src.setMinimumLoopTemperature(2.0)

        src_sp = make_const_schedule(model, f"Уставка_источника_ТН_{tag}", 12.0)
        src_mgr = openstudio.model.SetpointManagerScheduled(model, "Temperature", src_sp)
        src_mgr.addToNode(src.supplyOutletNode())

        src_pump = openstudio.model.PumpConstantSpeed(model)
        src_pump.setName(f"Насос_источника_ТН_{tag}")
        src_pump.setRatedPumpHead(30000.0)
        src_pump.setMotorEfficiency(0.85)
        src_pump.addToNode(src.supplyInletNode())

        src_comp = openstudio.model.PlantComponentTemperatureSource(model)
        src_comp.setName(f"Источник_тепла_ТН_{tag}")
        src_comp.setSourceTemperatureSchedule(src_sp)
        src_comp.autosizeDesignVolumeFlowRate()
        src.addSupplyBranchForComponent(src_comp)

        pl.addSupplyBranchForComponent(hp)
        src.addDemandBranchForComponent(hp)
        hp.setCondenserType("WaterSource")

        # Резерв источника, чтобы исключить срыв регулирования калорифера
        # в часы пиковой нагрузки/дефроста ТН.
        dh_backup = openstudio.model.DistrictHeating(model)
        dh_backup.setName(f"ИТП_Резерв_Теплосеть_{tag}")
        dh_backup.autosizeNominalCapacity()
        pl.addSupplyBranchForComponent(dh_backup)
    else:
        dh = openstudio.model.DistrictHeating(model)
        dh.setName(f"ИТП_Теплосеть_{tag}")
        dh.autosizeNominalCapacity()
        pl.addSupplyBranchForComponent(dh)

    pl.addDemandBranchForComponent(htg_coil)

    # ── 13. ПЕРИОД РАСЧЁТНОГО ДНЯ (SIZING PERIODS) ────────────────────────────
    # Зима: расчётный день СПб
    dd_win = openstudio.model.DesignDay(model)
    dd_win.setName("СПб_Зима_-24C")
    dd_win.setDayType("WinterDesignDay")
    dd_win.setMonth(1)
    dd_win.setDayOfMonth(21)
    dd_win.setMaximumDryBulbTemperature(-24.0)
    dd_win.setDailyDryBulbTemperatureRange(0.0)
    dd_win.setHumidityConditionType("Wetbulb")
    dd_win.setWetBulbOrDewPointAtMaximumDryBulb(-24.0)
    dd_win.setBarometricPressure(101325.0)
    dd_win.setWindSpeed(3.0)
    dd_win.setWindDirection(270.0)
    dd_win.setSolarModelIndicator("ASHRAEClearSky")
    dd_win.setSkyClearness(0.0)

    # Лето: параметр "А" (22°C)
    dd_sum = openstudio.model.DesignDay(model)
    dd_sum.setName("СПб_Лето_22C")
    dd_sum.setDayType("SummerDesignDay")
    dd_sum.setMonth(7)
    dd_sum.setDayOfMonth(21)
    dd_sum.setMaximumDryBulbTemperature(22.0)
    dd_sum.setDailyDryBulbTemperatureRange(8.5)
    dd_sum.setHumidityConditionType("Wetbulb")
    dd_sum.setWetBulbOrDewPointAtMaximumDryBulb(18.0)
    dd_sum.setBarometricPressure(101325.0)
    dd_sum.setWindSpeed(1.0)
    dd_sum.setWindDirection(180.0)
    dd_sum.setSolarModelIndicator("ASHRAEClearSky")
    dd_sum.setSkyClearness(1.0)

    # ── 14. ВЫХОДНЫЕ ПЕРЕМЕННЫЕ ───────────────────────────────────────────────
    outs = [
        ("Zone Mechanical Ventilation Current Density Volume Flow Rate", "Hourly"),
        ("Zone Mechanical Ventilation Mass Flow Rate", "Hourly"),
        ("Zone Air CO2 Concentration", "Hourly"),
        ("Zone People Occupant Count", "Hourly"),
        ("Zone Mean Air Temperature", "Hourly"),
        ("Fan Electric Energy", "Monthly"),
        ("Heating Coil Heating Energy", "Monthly"),
        ("District Heating Energy", "Monthly"),
        ("Zone Electric Equipment Electric Energy", "Monthly"),
        ("Zone Lights Electric Energy", "Monthly"),
        ("Zone People Sensible Heating Energy", "Monthly"),
        ("Air System Fan Electric Energy", "Monthly"),
        ("Zone Infiltration Current Density Volume Flow Rate", "Hourly"),
    ]
    for var, freq in outs:
        ov = openstudio.model.OutputVariable(var, model)
        ov.setReportingFrequency(freq)

    for m_name in ["Electricity:Facility", "DistrictHeating:Facility",
                   "Fans:Electricity", "Heating:DistrictHeating"]:
        om = openstudio.model.OutputMeter(model)
        om.setName(m_name)
        om.setReportingFrequency("Monthly")

    # ── 15. СОХРАНЕНИЕ ────────────────────────────────────────────────────────
    os.makedirs(MODELS_DIR, exist_ok=True)
    out_path = os.path.join(MODELS_DIR, f"kinozal_{tag}.osm")
    ok = model.save(openstudio.path(out_path), True)
    if ok:
        print(f"  [OK] Модель сохранена: {out_path}")
    else:
        print(f"  [!!] Ошибка сохранения: {out_path}")

    sim_dir = os.path.join(ROOT_DIR, f"sim_{tag}")
    os.makedirs(sim_dir, exist_ok=True)
    workflow_path = os.path.join(sim_dir, "workflow.osw")
    workflow = {
        "seed_file": out_path,
        "weather_file": WEATHER_PATH,
        "measure_paths": [],
        "steps": []
    }
    with open(workflow_path, "w", encoding="utf-8") as wf:
        json.dump(workflow, wf, ensure_ascii=False, indent=2)
    print(f"  [OK] Workflow: {workflow_path}")
    return model, out_path

# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 65)
    print("  СОЗДАНИЕ OSM-МОДЕЛЕЙ КИНОЗАЛА «ПЕТЕРБУРГ-КИНО»")
    print("  Большой зал 135 мест | 1-й этаж | Ярославский пр., 55А")
    print("=" * 65)

    print("\n▶ [1/6] Base — фиксированный расход 4050 м³/ч, рекуператор η=80%")
    m_base, p_base = build_model(dcv=False)

    print("\n▶ [2/6] DCV — управление по CO₂ (уставка 700 ppm), рекуп. η=80%")
    m_dcv, p_dcv = build_model(dcv=True)

    print("\n▶ [3/6] Base_HP — фиксированный расход, тепловой насос, рекуп. η=80%")
    m_base_hp, p_base_hp = build_model(dcv=False, heat_pump=True)

    print("\n▶ [4/6] DCV_HP — DCV (уставка 700 ppm) + тепловой насос, рекуп. η=80%")
    m_dcv_hp, p_dcv_hp = build_model(dcv=True, heat_pump=True)

    print("\n▶ [5/6] Base_Rec — фиксированный расход, высокоэффективный рекуператор η=90%")
    m_base_rec, p_base_rec = build_model(dcv=False, recuperator_high=True)

    print("\n▶ [6/6] DCV_Rec — DCV (уставка 700 ppm) + высокоэффективный рекуп. η=90%")
    m_dcv_rec, p_dcv_rec = build_model(dcv=True, recuperator_high=True)

    print("\n" + "=" * 65)
    print("  ГОТОВО — 6 МОДЕЛЕЙ")
    print(f"  Base:     {p_base}")
    print(f"  DCV:      {p_dcv}")
    print(f"  Base_HP:  {p_base_hp}")
    print(f"  DCV_HP:   {p_dcv_hp}")
    print(f"  Base_Rec: {p_base_rec}")
    print(f"  DCV_Rec:  {p_dcv_rec}")
    print("=" * 65)
