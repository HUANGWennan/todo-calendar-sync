# Microsoft To Do → Calendar Sync

[中文](#chinese) · [English](#english)

<a name="chinese"></a>
## 中文

通过 Microsoft Graph 将未完成、带截止日期的 To Do 任务单向镜像到同一账户的独立 Outlook 日历。将该账户加入 macOS「日历」即可查看。脚本可以在 macOS 的 PowerShell 或其他具备 PowerShell 7 的主机上运行，一套状态只由一个实例维护。

## 设置

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
$env:TODO_SYNC_ACCOUNT = 'you@example.com'
./sync.ps1 -Authenticate -DryRun
./sync.ps1 -DryRun
```

登录时采用 Microsoft 官方设备认证，委托权限为 `Tasks.Read`、`Calendars.ReadWrite`。核对 dry run 后，去掉 `-DryRun` 才会写入日历。默认登录个人账户，组织账户可用 `-TenantId organizations`，实际授权取决于组织策略。可用 `-SourceTimeZone`、`-CalendarTimeZone` 指定两端用于计算截止日期和全天事件的时区，默认均为 UTC。可用 `-CalendarName` 指定独立日历名。

脚本先完整读入所有源任务，再写入日历。它用稳定的任务标识管理自己创建的事件。完成或移除源任务时，脚本会删除对应的镜像事件。手动创建的普通事件保持原样。重复事件、账户不一致或同步目标缺失时，脚本会停下，等待检查。

同步状态默认放在本机用户目录下的 `.todo-calendar-sync`。定期运行时复用同一状态目录，并用主机自己的计划任务启动。

---

<a name="english"></a>
## English

Mirrors unfinished Microsoft To Do tasks with due dates into a dedicated Outlook calendar in the same account through Microsoft Graph. Add the account to macOS Calendar to see the mirrored events. The script can run on macOS or another PowerShell 7 host. Use only one instance per state directory.

### Set up

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
$env:TODO_SYNC_ACCOUNT = 'you@example.com'
./sync.ps1 -Authenticate -DryRun
./sync.ps1 -DryRun
```

Sign in through Microsoft's device-code flow. The script requests delegated `Tasks.Read` and `Calendars.ReadWrite` permissions. Review the dry run before running without `-DryRun`. Personal accounts are the default; for an organization, try `-TenantId organizations` if organizational policy allows it. Use `-SourceTimeZone` and `-CalendarTimeZone` to control due dates and all-day events (both default to UTC), and `-CalendarName` to name the destination calendar.

The script reads the full source list before writing and uses stable task identifiers to track its own events. Completing or removing a task removes its mirror. Manually created calendar events are left alone. If there are duplicate events, an account mismatch, or no destination calendar, the sync stops for review.

State defaults to `.todo-calendar-sync` in the home directory. For periodic runs, reuse one state directory and use the host's scheduler.
