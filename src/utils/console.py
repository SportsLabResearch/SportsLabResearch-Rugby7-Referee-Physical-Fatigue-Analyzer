from src.config.app_info import (
    APP_SHORT_NAME,
    APP_VERSION,
    ORGANIZATION,
)


def show_banner():
    print("=" * 60)
    print(ORGANIZATION)
    print(f"{APP_SHORT_NAME} {APP_VERSION}")
    print("=" * 60)


def show_main_menu():
    print()
    print("1. Analizar un árbitro")
    print("2. Analizar varios árbitros")
    print("3. Analizar todos los árbitros")
    print("4. Generar versión pública (anonimizada)")
    print("0. Salir")
    print("-" * 60)
