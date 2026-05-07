# TaskFlow Role Access Matrix

## Task Access

| Capability | Owner | Manager | Member |
| --- | --- | --- | --- |
| View tasks | All tasks | Tasks in own team, tasks created by them, tasks owned by them | Tasks owned by them or created by them |
| Create task | Any active member | Own team members, other managers, owner for escalation | Self only |
| Edit task fields | Any task | Tasks in accessible scope | Own/created tasks only where allowed by UI |
| Route task | Any active member | Own team members, other managers, owner for escalation | Not allowed |
| Change status | Any non-recurring task | Accessible non-recurring tasks | Own/created non-recurring tasks where transition is allowed |
| Complete task | Any allowed task | Accessible allowed tasks | Own/created allowed tasks |
| Archive task | Any task | Accessible tasks | Not allowed |
| Reverse last status | Owner only | Not allowed | Not allowed |
| Comments | Accessible tasks | Accessible tasks | Own/created tasks |
| Checklists/subtasks | Accessible tasks | Accessible tasks | Own/created tasks |

## HOD Desk Access

| Capability | Owner | Manager | Member |
| --- | --- | --- | --- |
| View HOD Desk | All HOD/private spaces | Invited/own manager spaces | Hidden from navigation |
| Create HOD space | Yes | Yes, for private manager coordination | No |
| Invite managers | Any manager | Managers only, owner included automatically where applicable | No |
| Escalate task to HOD | Any task | Accessible tasks | No; member escalates to manager first |
| Approve/Delete escalation messages | Yes | Space participants/managers where allowed | No |
| Open linked task from HOD | Any linked task | Linked tasks in accessible scope | No |
| HOD comments/messages | Any HOD space | Spaces they belong to | No |

## Operating Rules

- Members should work inside task drawer comments, not HOD Desk.
- HOD Desk is for manager/owner coordination and cross-department escalation.
- Assignment rules should stay symmetric across create, edit, route, bulk route, and comments.
- Recurring tasks remain special: they are shown in `On Hold`, should not show overdue, and should only send reminders while active/on hold.
