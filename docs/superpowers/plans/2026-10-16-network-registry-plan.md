# Адреса заявки и Superset registry plan

Base2f2d1ab/PR20, inline. Готовые14 checks не переписываются. Цель: общий индекс
применяется к эффективным адресам, создаются immutable entries будущей сводки
и независимый NET.IP_REGISTRY; ApplyRegistry добавляет15-й исход к готовому14.

1. ATNetworkRegistryPlanner.Run(source, compiledIndex) возвращает
   ATNetworkRegistryResult (State/Entries/Report). Верхние SOURCE/DEST и нижние
   ADDRESS разбираются по объявлениям прежним IP parser с compact-prefix context.
   Domain math key дедуплицирует одинаковый IP, /32 и эквивалентную сеть/маску,
   без расширения маски в отдельные entries. Все refs сохраняются. Порядок IDs —
   первое появление сверху, затем снизу. Группы не раскрываются и не создают
   бесполезных предупреждений; отдельные числовые нижние адреса участвуют всегда.
   Неверный numeric fragment/Excel error явно incomplete; корректные соседние
   объявления сохраняются, compact-prefix сбрасывается после ошибки/группы.
2. ATNetworkRegistryEntry хранит Id/Address/Key/References/Hostnames/Matches/
   Status/SourceComplete/HostnameRefined/Domain. После Finish frozen. Ссылки
   сохраняют FIELD_ID и CELL_REF. Один запрос на уникальный math key.
   Hostname нижней строки привязан только к тому же math key, без угадывания.
   При точном hostname среди IP hits разрешено сужение до всех точных совпадений.
   При отсутствии точного совпадения IP hits остаются. Ошибка source hostname
   блокирует сужение, сохраняя IP hits и manual warning.
3. FOUND/NOT_FOUND/UNCONFIRMED сохраняются для каждой записи, включая отсутствие
   совпадений. NOT_FOUND возможен только при полном источнике. Нет index/он
   повреждён/SourceComplete=False — явная неполнота, не false not-found.
   Published Rulebook: REGISTRY_NOT_FOUND WARNING, не user-input error,
   key NET.IP_REGISTRY|WARNING|Проверка IP, fallback coordinator/Superset-only.
   Outcome SKIPPED при отсутствии IP/неполноте/не найденном адресе; ERROR override
   имеет приоритет благодаря Incomplete flag. Бюджет/структурная ошибка — BLOCKED
   с пустыми entries, без partial success. Общий budget200000,10000 entries,
  100000 aggregate links, report990 issues, messages32767.
4. ATNetworkAccessChecks.ApplyRegistry(previous, registry, catalog) независимо
   оформляет новый check и объединяет15 outcomes/decisions; старые сообщения не
   переоформляются. Новый presentation limit не теряет14 старых исходов.
   Excel report отличает результат IP registry от ещё не созданной итоговой
   Superset copy/координаторской сводки. Полный writer/UI — следующий этап.
5. TDD: группа+нижниеIP+unmatched; duplicates upper/lower/IPvs/32; compact/mask/
   CIDR and invalid fragments; hostname exact/fallback/ties/source error;
   source incomplete/unavailable; freeze/bounds; combined15/writer/routes and
   severity. Full Excel90s (текущая среда), actual blank referencesReadOnly,
   independent review, static7/gate9/clean; Draft PR over20, no merge/release.

Review focus: all lower IP preserved with groups; no lost refs/no duplicate queries;
partial matches never hidden by missing hostname; incomplete source not definitive
absence; independent result/presentation failure preserves14; no partial entries
on budget/system errors. Rulebook hostname/service application, actual-file
freshness and final coordinator workbook remain pending.
