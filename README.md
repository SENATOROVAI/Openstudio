# Openstudio: 6 моделей вентиляции кинозала

## Краткое описание
Проект содержит воспроизводимый расчётный контур OpenStudio/EnergyPlus для 6 сценариев:
- `Base`
- `DCV`
- `Base_HP`
- `DCV_HP`
- `Base_Rec`
- `DCV_Rec`

Результат проекта: единый валидированный отчёт DOCX с таблицами, графиками и выводами, синхронизированными с фактическими SQL-результатами всех 6 моделей.

## Структура
- `models/` — OSM-файлы 6 моделей.
- `sim_*/workflow.osw` — workflow для запуска каждой модели.
- `scripts/create_kinozal_model.py` — генерация моделей и workflow.
- `scripts/extract_results.py` — извлечение метрик из `eplusout.sql` в `results/results.json`.
- `scripts/generate_report.py` — сборка итогового DOCX и графиков.
- `results/results.json` — агрегированные данные по 6 моделям.
- `results/figures/` — рисунки для отчёта.
- `Отчёт_Кинозал_DCV_OpenStudio 3.docx` — итоговый отчёт.

## Полный запуск (пошагово)
```bash
cd /Users/m/Documents/Openstudio

# 1) Пересоздать модели и workflow
/usr/bin/python3 scripts/create_kinozal_model.py

# 2) Запустить все 6 симуляций
openstudio run -w sim_Base/workflow.osw
openstudio run -w sim_DCV/workflow.osw
openstudio run -w sim_Base_HP/workflow.osw
openstudio run -w sim_DCV_HP/workflow.osw
openstudio run -w sim_Base_Rec/workflow.osw
openstudio run -w sim_DCV_Rec/workflow.osw

# 3) Собрать численные результаты
/usr/bin/python3 scripts/extract_results.py

# 4) Сформировать отчет с таблицами и картинками
/usr/bin/python3 scripts/generate_report.py
```

## Контроль валидности
```bash
cd /Users/m/Documents/Openstudio
for s in Base DCV Base_HP DCV_HP Base_Rec DCV_Rec; do
  echo "--- sim_$s"
  rg "EnergyPlus Completed Successfully|Severe Errors" sim_${s}/run/eplusout.err
 done
```
Ожидаемое состояние: для каждого сценария `0 Severe Errors`.

## Краткие выводы (по текущему results.json)
- `DCV` относительно `Base`: снижение общего электропотребления около `8.16%`.
- Снижение электропотребления вентиляторов в паре `DCV/Base`: около `74.49%`.
- `DCV_HP` относительно `Base_HP`: снижение общего электропотребления около `8.18%`.
- Все 6 сценариев завершены без `Severe Errors`; данные в DOCX формируются из актуального `results/results.json`.
