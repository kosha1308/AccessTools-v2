# Этап 2B.2 — правила по строкам таблиц

Пользователь делегировал технические решения. Цель первого рабочего выпуска
сохраняется: NETWORK_ACCESS, Superset и сводка координатора. Сейчас выполняется
один следующий шаг — корректное вычисление по строкам без чтения реальной книги.

Выбран обход уже типизированного AST. Повторное разрешение имён в Runtime или
flatten строк в document fields отклонены: это теряет lexical aliases и корреляцию.
Компилятор сохраняет FieldId/TableId/RowAlias и NamedScope. Serializer не меняется.
Unqualified row field после compile имеет точный alias; Runtime ничего не угадывает.

ATRuntimeRow хранит TableId, устойчивый RowId и типизированные результаты полей,
при необходимости адрес ячейки как metadata. Snapshot замораживает состав полей;
getters защищают значения. COLLECTION<ROW<T>> хранит отдельную копию списка
неизменяемых rows; тип таблицы и уникальность RowId обязательны. Пустая коллекция
является PRESENT с Count=0. Payload недоступен через Scalar/Item(list).

ALLROWS/ANYROW/NOROWS/FILTER исполняют predicate слева направо по строкам.
Вложенные aliases видят внешние строки; текущая строка восстанавливается после
каждого вложенного обхода. DOCUMENT named definitions и defaults получают пустой
row scope, ROW определения — caller row scope; supplied arguments всегда caller.
DOCUMENT definition может объявить alias с таким же именем, как у caller, без захвата.

Первый non-OK predicate распространяется без изменения состояния/первичной
диагностики. Boolean predicates требуют PRESENT BOOL. После decisive Boolean
остальные строки не вычисляются и получают NOT_EVALUATED trace с RowId.
ALLROWS(empty)=TRUE, ANYROW(empty)=FALSE, NOROWS(empty)=TRUE. FILTER сохраняет
порядок/RowId/тип таблицы; при ошибке частичный результат не выдаётся.
COUNT возвращает размер коллекции, EXISTS — Count>0; поля строк не читаются ради
COUNT/EXISTS. DISTINCT_COUNT и collection equality остаются unsupported.

ATEvalBudget контролирует количество строк MaxListItems до копирования, шаг на
каждую обработанную/пропущенную строку, глубину и общий trace. Неиспользованные поля
не клонируются/не валидируются. Отдельное поле проверяется до возврата его копии.
Trace расширяется RowId/TableId/FieldId/CellRef без raw payload; существующие spans
и SourceId сохраняются. Metadata не исполняется как адрес/формула. При лимите trace
первичный non-OK сохраняется; успешный результат заменяется ERROR с лимитом.

Приёмка: aliases и scope/незахват, correlated IP predicate, empty truth table,
lazy skipping с плохим полем в следующей строке, все non-OK states, фильтрация,
ошибки типов/полей, immutability, бюджеты и восстановление после ошибок. Существующие
9563 assertions сохраняются. Excel compile/full tests, independent review,
tooling/gate, clean clone и CI. Excel adapter/merged cells/Superset/GUI — следующий шаг.
