#!/usr/bin/env python3
"""
Generatore di presentazioni con Gemini (client google-genai)
=============================================================
Utilizza il pacchetto google-genai (non google-generativeai).
"""

import os
import sys
import datetime
from google import genai
from google.genai.types import GenerateContentConfig

# ----------------------------------------------------------------------
# 1. Gestione API Key
# ----------------------------------------------------------------------
def get_api_key():
    """Recupera la chiave API dalla variabile d'ambiente SLIDES_API_KEY,
       altrimenti la chiede all'utente."""
    key = os.environ.get("SLIDES_API_KEY")
    if key:
        return key
    return input("Inserisci la tua API key di Google Gemini: ").strip()

# ----------------------------------------------------------------------
# 2. Lettura del template di prompt
# ----------------------------------------------------------------------
def read_prompt_template(file_path="prompt.txt"):
    """Legge il contenuto del file 'prompt.txt' che contiene le regole
       di strutturazione della presentazione."""
    if not os.path.exists(file_path):
        print(f"❌ Errore: il file '{file_path}' non esiste nella directory corrente.")
        sys.exit(1)
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read()

# ----------------------------------------------------------------------
# 3. Generazione del contenuto tramite Gemini (client)
# ----------------------------------------------------------------------
def generate_slides(topic, num_slides, prompt_template,
                    model_name="gemma-4-31b-it"):
    """
    Compone il prompt finale e invoca il modello Gemini per generare il Markdown.
    """
    # Prompt finale: regole + richiesta specifica
    full_prompt = (
        f"{prompt_template}\n\n"
        f"Ora, genera una presentazione in Markdown sull'argomento '{topic}' "
        f"con esattamente {num_slides} slide normali (cioè con intestazioni ###). "
        "La presentazione deve avere un titolo principale (#), almeno una sezione (##) "
        "e il numero richiesto di slide (###). Segui rigorosamente le regole sopra."
    )

    # Crea il client con la chiave API
    api_key = get_api_key()
    client = genai.Client(api_key=api_key)

    try:
        response = client.models.generate_content(
            model=model_name,
            contents=full_prompt,
            config=GenerateContentConfig(
                temperature=0.7,
                max_output_tokens=8192,
            )
        )
        return response.text
    except Exception as e:
        print(f"❌ Errore durante la generazione: {e}")
        sys.exit(1)

# ----------------------------------------------------------------------
# 4. Salvataggio del risultato
# ----------------------------------------------------------------------
def save_output(content):
    """Salva il contenuto in un file con timestamp nel nome."""
    now = datetime.datetime.now()
    filename = now.strftime("%Y%m%d%H%M") + ".md"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"✅ Contenuto salvato in {filename}")

# ----------------------------------------------------------------------
# 5. Main
# ----------------------------------------------------------------------
def main():
    print("Benvenuto nel generatore di presentazioni con Gemini\n")

    # Richiesta argomento
    topic = input("Inserisci l'argomento della presentazione: ").strip()
    if not topic:
        print("❌ Argomento non valido.")
        sys.exit(1)

    # Richiesta numero di slide normali (H3)
    num_str = input("Inserisci il numero di slide normali (###) desiderato: ").strip()
    try:
        num_slides = int(num_str)
        if num_slides <= 0:
            raise ValueError
    except ValueError:
        print("❌ Inserisci un numero intero positivo.")
        sys.exit(1)

    # Lettura del template
    prompt_template = read_prompt_template()

    # Generazione
    print(f"\n⏳ Generazione della presentazione su '{topic}' con {num_slides} slide...")
    content = generate_slides(topic, num_slides, prompt_template)

    # Salvataggio
    save_output(content)
    print("🎉 Operazione completata.")

if __name__ == "__main__":
    main()