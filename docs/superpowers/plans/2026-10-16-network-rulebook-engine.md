# Вычисление общих правил NETWORK_ACCESS

Продолжение PR17 (9027ece). Цель этапа: неизменяемый вычислитель таблицы RULES,
который сможет использовать проверка заявки и будущий Superset/coordinator runner.
Тринадцать готовых проверок и их presentation pass не меняются.

## Контракт и решения

- Compile принимает RAW RULES и опубликованный каталог протоколов. Пустой ACTIVE
  означает неактивность согласно комментарию реальной книги; неизвестные флаги,
  повреждённые активные условия и неподдерживаемые сочетания блокируют всю компиляцию.
- IP — пересечение для срабатывания правила, не разрешение инициатору. Это другой
  смысл, чем полный subset в персональной policy. Порты — только полное покрытие
  одним правилом; объединение разных правил без бизнес-соглашения не вводится.
- Hostname exact > prefix > contains, затем длина совпавшего шаблона. Регистр и
  конечная точка незначимы. Все равные лучшие правила сохраняются; правила без
  hostname применяются независимо. Порядок строк не выбирает победителя.
- Technical: источник/назначение, набор портов, один протокол. Hostname-условия
  при подходящих остальных условиях требуют ручной проверки. Service-only правила
  принадлежат Superset-контексту. Superset: IP/hostname/существующий сервис;
  port/protocol правила здесь не вычисляются. Service+port/protocol запрещено,
  как в опубликованной прежней схеме.
- Результат — immutable COLLECTION<ROW<RULE_MATCH>>, ссылки исходных правил,
  все семь outputs отдельно: name/service/group/approval/comment/severity/instruction.
  Конфликт разных service/group сохраняется флагом, без случайного выбора.
- Нет условий/нет результата — ошибка активной строки. Instruction считается
  самостоятельным полезным результатом. Пустая severity = INFO (legacy default).
- Общий budget на compile/evaluation, без разворачивания IP/портов. Ошибка/лимит
  не выдаёт частичные результаты. Исходные книги не изменяются.

## Выполнение (inline)

1. Тесты компиляции RAW, inactive/defaults, malformed/Excel errors, full ports,
   scoped IP, specificity/ties/conflicts, service context, unknown input и budgets.
   Новые ATNetworkRulebookCompiler/Result, ATNetworkRulebook/Rule и text helper.
2. Запустить RED в настоящем Excel, реализовать, получить GREEN всей регрессии.
3. Один свежий обзор ветки, исправления с RED/GREEN; tooling7/7, gate9/9,
   чистая копия с 60s. Обновить статус; draft PR поверх PR17, без merge/release.

## Review focus

Частичный порт не совпадает; неизвестное значение не даёт false success;
hostname specificity считается только среди подходящих base conditions;
исходные ref и инструкции не теряются; budgets и mutable collection не обходят
immutable модель. Подключение к 13-check runner, Superset matcher и итоговая книга
остаются следующими задачами после этого вычислителя.
