# VASPars — автодеплой и запуск

Проект парсит скриншоты измерений ВАЦ и сохраняет результаты в `result.xlsx`.

## Что устанавливает автодеплой
Скрипт `deploy.ps1` автоматически:
1. Проверяет наличие Python.
2. Создаёт виртуальное окружение `.venv`.
3. Устанавливает Python-зависимости из `requirements.txt`.
4. Проверяет наличие Tesseract OCR.
5. Создаёт `run_parser.bat` для удобного запуска.

## Требования
- Windows 10/11
- Python 3.10+
- PowerShell 5+
- (Желательно) установленный Tesseract OCR

## Быстрый старт на рабочем ПК
1. Скопируйте папку проекта на рабочий компьютер.
2. Откройте PowerShell в папке проекта.
3. Выполните:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\deploy.ps1
   ```
4. После завершения запустите:
   ```bat
   run_parser.bat
   ```

## Установка Tesseract (если не найден)
Если в конце деплоя появится предупреждение, установите Tesseract:
```powershell
winget install UB-Mannheim.TesseractOCR
```

После установки запустите `deploy.ps1` ещё раз.

## Обновление зависимостей
Если проект обновился, выполните заново:
```powershell
powershell -ExecutionPolicy Bypass -File .\deploy.ps1
```

## Структура файлов деплоя
- `deploy.ps1` — автодеплой
- `requirements.txt` — список Python-пакетов
- `run_parser.bat` — запуск приложения в настроенном окружении
