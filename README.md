# Validador de Glebas

Aplicativo desktop em Python com `CustomTkinter` para validar planilhas de glebas e apontar inconsistencias de coordenadas e sequenciamento.

## Requisitos

- Python 3.10
- Dependencias do projeto instaladas na `.venv`

## Como executar

```powershell
.\.venv\Scripts\python.exe .\validador_glebas_app.py
```

## Como gerar o executavel

```powershell
.\build_executavel.bat
```

## Arquivos principais

- `validador_glebas_app.py`: interface desktop.
- `validador2_glebas.py`: versao auxiliar da logica.
- `build_executavel.bat`: gera o `.exe` com PyInstaller.
- `assets/validador_glebas_icon.ico`: icone do aplicativo.

## Exemplos

O repositorio inclui planilhas `.xls` de exemplo para teste rapido do validador.
