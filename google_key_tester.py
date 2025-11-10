# google_key_tester.py
# ------------------------------------------------------
# Testa se sua GOOGLE_API_KEY está válida e se as APIs
# principais (Geocoding, Directions, Distance Matrix)
# estão ativas.
# ------------------------------------------------------

import requests

def testar_google_api_key(api_key: str):
    """
    Testa a chave do Google Maps em três endpoints:
    Geocoding, Directions e Distance Matrix.
    Retorna dicionário com status e mensagens.
    """

    resultados = {}

    # 1️⃣ Teste GEOCODING
    try:
        url_geo = f"https://maps.googleapis.com/maps/api/geocode/json?address=Pirassununga&key={api_key}"
        r = requests.get(url_geo, timeout=10)
        dados = r.json()
        resultados["Geocoding"] = dados.get("status", "SEM RESPOSTA")
    except Exception as e:
        resultados["Geocoding"] = f"ERRO: {e}"

    # 2️⃣ Teste DIRECTIONS
    try:
        origem = "Pirassununga"
        destino = "Leme"
        url_dir = f"https://maps.googleapis.com/maps/api/directions/json?origin={origem}&destination={destino}&key={api_key}"
        r = requests.get(url_dir, timeout=10)
        dados = r.json()
        resultados["Directions"] = dados.get("status", "SEM RESPOSTA")
    except Exception as e:
        resultados["Directions"] = f"ERRO: {e}"

    # 3️⃣ Teste DISTANCE MATRIX
    try:
        origem = "Pirassununga"
        destino = "Leme"
        url_dm = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={origem}&destinations={destino}&key={api_key}"
        r = requests.get(url_dm, timeout=10)
        dados = r.json()
        resultados["DistanceMatrix"] = dados.get("status", "SEM RESPOSTA")
    except Exception as e:
        resultados["DistanceMatrix"] = f"ERRO: {e}"

    return resultados


if __name__ == "__main__":
    print("🔍 Teste de Chave Google API")
    chave = input("Digite ou cole sua GOOGLE_API_KEY: ").strip()

    if not chave:
        print("⚠️ Nenhuma chave informada.")
    else:
        resultado = testar_google_api_key(chave)
        print("\n📊 Resultados:")
        for api, status in resultado.items():
            print(f" - {api}: {status}")

        if all(s == "OK" for s in resultado.values()):
            print("\n✅ Todas as APIs respondendo normalmente!")
        else:
            print("\n⚠️ Alguma API retornou erro ou precisa ser habilitada no console Google.")
