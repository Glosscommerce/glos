# MEMOIRE CORRECTEUR PRO gloss — Kit d’installation Windows sy fampiasana amin’ny téléphone

## Inona no ao anatin’ity kit ity?
- **INSTALLER_WINDOWS.bat**: mametraka ilay application ao amin’ny Windows, mametraka ireo bibliothèques Python, ary mamorona raccourci eo amin’ny Bureau.
- **LANCER_APP.vbs**: manokatra ilay application amin’ny fomba tsy mampiseho CMD.
- **LANCER_WINDOWS.bat**: fomba fanombohana hita maso, ilaina raha misy erreur.
- **LANCER_TELEPHONE_WIFI.bat**: fomba fampiasana amin’ny smartphone mifandray amin’ilay Wi‑Fi mitovy amin’ny PC.
- **app/app.py**: kaody an’ilay application.

## A. Fametrahana amin’ny Windows
1. Ataovy **Extract All / Extraire tout** amin’ilay ZIP; aza alefa mivantana avy ao amin’ny WinRAR na dossier `Temp`.
2. Mila Python 3.11 na vaovao kokoa sy Internet amin’ny fametrahana voalohany. Python: https://www.python.org/downloads/windows/ . Raha mametraka ianao dia asio marika **Add Python to PATH**.
3. Sokafy ilay dossier nalaina, dia tsindrio indroa **INSTALLER_WINDOWS.bat**.
4. Rehefa vita, tokony hisy **MEMOIRE CORRECTEUR PRO gloss** eo amin’ny Bureau. Tsindrio indroa hanokafana.
5. Mihazakazaka any ambadika ilay serveur; noho izany tsy ilaina ny mitazona CMD misokatra eo amin’ny efijery. Raha nakatona ny session na naverina natomboka ny PC dia tsindrio indray ilay raccourci.

## B. Fampiasana amin’ny smartphone raha eo amin’ny Wi‑Fi iray ihany
Mila mandeha sy tsy maty ilay PC, ary mifandray amin’ny **Wi‑Fi mitovy** ny téléphone sy PC.
1. Sokafy ny dossier napetraka: `%LOCALAPPDATA%\Programs\MemoireCorrecteurProGloss`.
2. Alefaso ny **LANCER_TELEPHONE_WIFI.bat**.
3. Ao amin’ny CMD hafa, soraty `ipconfig`, dia tadiavo ny **IPv4 Address** an’ny Wi‑Fi na Ethernet ampiasain’ny PC (ohatra `192.168.1.23`).
4. Ao amin’ny Chrome amin’ny téléphone dia soraty `http://ADRESSE-IP:8502` (soloina ilay tena IP, ohatra `http://192.168.1.23:8502`).
5. Raha manontany ny Windows Firewall dia avelao amin’ny **Private networks** ihany. Aza alefa amin’ny réseau public.
6. Raha te-hametraka kisary eo amin’ny écran d’accueil: Chrome Android → menu ⋮ → **Ajouter à l’écran d’accueil** na **Installer l’application** raha miseho.
7. Aza akatona ilay CMD amin’ny PC mandritra ny fampiasana amin’ny téléphone.

Ity fomba Wi‑Fi ity dia tsy mandeha lavitra: tsy maintsy mandeha ny PC sy ilay serveur ary ao amin’ny réseau mitovy ny téléphone.

## C. Fampiasana amin’ny téléphone na aiza na aiza
Mba hampiasaina rehefa tsy eo akaikin’ny PC, dia mila **mametraka (héberger) an’ilay app amin’ny Internet**. Ity kit ity dia efa natao ho an’ny Streamlit Community Cloud:
1. Mamoròna kaonty GitHub, ary repository iray. Ampidiro ao amin’ny repository ny atiny ao amin’ity dossier ity, ankoatra `.venv` raha efa misy.
2. Sokafy https://share.streamlit.io/ ary ampifandraiso amin’ny GitHub.
3. Safidio **Create app**, repository-nao, branch `main`, ary main file path `app/app.py`, dia **Deploy**. Torolalana ofisialy: https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy
4. Alohan’ny hizarana ny rohy, ao amin’ny **Advanced settings → Secrets**, apetraho ohatra izao, ary ovao ho teny miafina matanjaka manokana:

   ```toml
   APP_PASSWORD = "OVAO_ITY_HO_TENY_MIAFINA_MATANJAKA"
   ```

5. Rehefa vita ny deployment dia homena rohy `https://...streamlit.app` ianao. Sokafy amin’ny téléphone, midira amin’ny mot de passe, dia fidio Chrome → ⋮ → **Ajouter à l’écran d’accueil**.

**Zava-dehibe momba ny tsiambaratelo:** mémoire an’ny mpianatra dia mety ahitana angona manokana. Aza mampakatra antontan-taratasy tsiambaratelo amin’ny app an-tserasera raha tsy matoky sy mahalala ny fomba itehirizana sy fanodinana azy. Mametraha `APP_PASSWORD` alohan’ny hizarana app; aza apetraka ao amin’ny kaody na GitHub ny teny miafina, na ny clé API.

## Fetrany
- Ity dia app web Python/Streamlit, fa **tsy APK Android natively compilé** ary tsy `.exe` tokana mitondra Python ao anatiny. Ny installer Windows dia mametraka ny dépendances amin’ny PC, ka ilaina ny Python sy Internet amin’ny installation voalohany.
- Ny shortcut eo amin’ny smartphone dia mankany amin’ilay site web; mila Internet sy hébergement ny fampiasana lavitra. Tsy miasa offline amin’ny smartphone ity dikan-teny ity.
- Ny reformulation IA dia mila `OPENAI_API_KEY` sy mety hisy sarany API. Tsy ilaina amin’ny fanamarinana fototra sy mise en page.
- Ny famindrana ho APK Android tena izy dia dingana hafa: mila mamorona interface Android na wrapper ary manamboatra/build/sign APK aorian’ny fisian’ny rohy web raikitra.
