# OpenStudio/EnergyPlus: результаты моделирования (6 сценариев)

## Назначение
Папка содержит готовые результаты расчёта системы вентиляции кинозала для 6 сценариев:
- `Base`
- `DCV`
- `Base_HP`
- `DCV_HP`
- `Base_Rec`
- `DCV_Rec`

## Что входит в поставку
- `models/` — OSM-модели сценариев;
- `sim_*/workflow.osw` — workflow-файлы для запуска EnergyPlus/OpenStudio;
- `sim_*/run/` — результаты выполненных прогонов (`eplusout.sql`, `eplusout.err`, `eplustbl.htm` и др.);
- `results/results.json` — сводные численные показатели;
- `results/figures/*.png` — итоговые графики (версия для 6 моделей);
- `Отчёт_Кинозал_DCV_OpenStudio 3.docx` — итоговый отчёт.

## Требования для повторного запуска
- OpenStudio CLI (`openstudio`);
- доступ к файлу погоды `weather/Saint-Petersburg.epw`.

## Как повторно выполнить расчёт
```bash
cd /Users/m/Documents/Openstudio

openstudio run -w sim_Base/workflow.osw
openstudio run -w sim_DCV/workflow.osw
openstudio run -w sim_Base_HP/workflow.osw
openstudio run -w sim_DCV_HP/workflow.osw
openstudio run -w sim_Base_Rec/workflow.osw
openstudio run -w sim_DCV_Rec/workflow.osw
```

## Проверка корректности прогонов
```bash
cd /Users/m/Documents/Openstudio
for s in Base DCV Base_HP DCV_HP Base_Rec DCV_Rec; do
  echo "--- sim_$s"
  rg "EnergyPlus Completed Successfully|Severe Errors" sim_${s}/run/eplusout.err
done
```

Ожидаемый результат: во всех сценариях `0 Severe Errors`.

## Основные итоговые файлы
- `results/results.json`
- `results/figures/fig1_total_electricity_6models.png`
- `results/figures/fig2_fan_electricity_6models.png`
- `results/figures/fig3_monthly_fans_6models.png`
- `results/figures/fig4_co2_average_6models.png`
- `Отчёт_Кинозал_DCV_OpenStudio 3.docx`
