$NomeTarefa = "AtualizacaoPedidosNotion"
$PastaApp   = "C:\Users\l3ti\Ferramenta de Atualizacoes de Pedidos"
$Python     = (Get-Command python).Source

$Acao       = New-ScheduledTaskAction -Execute $Python -Argument "atualizar_pedidos.py" -WorkingDirectory $PastaApp
$Gatilho    = New-ScheduledTaskTrigger -Daily -At "00:00"
$Config     = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

Register-ScheduledTask -TaskName $NomeTarefa -Action $Acao -Trigger $Gatilho -Settings $Config -RunLevel Highest -Force
Write-Host "Tarefa agendada criada com sucesso! Roda todo dia a meia-noite."
