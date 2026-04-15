KSeF - Zestawienie faktur zakupu (wersja 1)

Co robi program:
- otwiera KSeF w Microsoft Edge,
- użytkownik loguje się ręcznie,
- program próbuje wpisać zakres dat,
- pobiera dane z tabeli faktur zakupu,
- zapisuje Excel w układzie podobnym do Twoich plików.

Ważne:
Ta wersja jest zrobiona jako solidny szkielet pod prawdziwy portal KSeF.
Bez podglądu konkretnego ekranu KSeF nie da się uczciwie wpisać idealnych selektorów na sztywno.
Dlatego w programie jest plik ksef_config.json, gdzie można łatwo dopracować selektory tabeli i przycisku następnej strony.

Jak uruchomić:
1. Zainstaluj Python 3.11 lub nowszy.
2. Otwórz CMD w folderze programu.
3. Wpisz:
   pip install -r requirements.txt
4. Uruchom:
   python ksef_zestawienie_gui.py

Jak używać:
1. Kliknij 'Otwórz KSeF'.
2. Zaloguj się ręcznie.
3. Wejdź w listę faktur zakupu.
4. Kliknij 'Ustaw daty w KSeF' albo ustaw daty ręcznie na stronie.
5. Kliknij 'Pobierz zestawienie'.
6. Program zapisze plik Excel.

Jeśli program nie złapie tabeli lub paginacji:
- otwórz ksef_config.json
- popraw sekcję table_selectors
- popraw sekcję next_button_xpaths

Docelowo po jednym teście na prawdziwym ekranie KSeF można to bardzo szybko doszlifować.
