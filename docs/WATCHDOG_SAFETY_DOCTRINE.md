# Watchdog Safety Doctrine

The Outlook watchdog follows these rules:

1. Inspect broadly.
2. Report clearly.
3. Modify narrowly.
4. Only modify scaffold-managed messages.
5. Never modify unmarked user mail.
6. Prefer `ACTION_REQUIRED` over destructive repair when uncertain.

The watchdog may inspect all Outbox mail for diagnostics, but repair actions are limited to messages with a valid managed email key owned by this scaffold.

