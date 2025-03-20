# AutoContract-VBA
O AutoContract é um aplicativo desenvolvido em VBA para a confecção de Minutas pré configuradas em Word ou PDF.

## Da Ineficiência à Automação
A confecção manual de contratos era um processo moroso, repetitivo e propenso a erros, causando perda de tempo e retrabalho. Antes da automação, os usuários tinham que copiar e colar informações da tela do ServiceNow, e para complementar os dados, precisavam consultar prints do sistema PCOM, enviados em uma etapa prévia, e também, copiar e colar manualmente cada informação no contrato. Esse método não só aumentava o risco de erros humanos, como também tornava cada contrato um trabalho demorado e cansativo, sujeito a inconsistências e falta de padronização. Pequenos erros podiam resultar em retrabalho, falhas contratuais e atrasos na entrega dos documentos.

Para eliminar essas dificuldades, o aplicativo automatiza completamente o processo. Agora, o usuário apenas faz o upload do arquivo .json do ServiceNow e seleciona as alterações desejadas. O aplicativo, então, busca automaticamente os dados complementares no sistema PCOM, estruturando todas as informações sem a necessidade de intervenção manual. Com um único clique, a minuta do contrato é gerada, completa, padronizada e pronta para revisão.

Essa automação eliminou o trabalho manual, reduziu drasticamente o tempo gasto na confecção de contratos e minimizou erros, garantindo mais eficiência, precisão e escalabilidade. O que antes era um processo burocrático e suscetível a falhas, agora é rápido, confiável e otimizado, permitindo que os profissionais se concentrem no que realmente importa: análise, revisão e decisões estratégicas.

## Funcionamento

### *Tela de Login*
Ao abrir o aplicativo AUTOCONTRACT.xlsm, o usuário é direcionado para uma tela de login, onde deve inserir corretamente seu número de usuário. Esse passo é fundamental para que o aplicativo realize a busca adequada pelos downloads no computador do usuário.

A tela de login utiliza uma senha padrão para todos os usuários, permitindo-lhes gerar minutas em modo somente leitura, sem a possibilidade de editar o aplicativo. Todos os usuários podem acessar o aplicativo e confeccionar minutas simultaneamente, sem a necessidade — e nem a recomendação — de fazer cópias do arquivo em suas máquinas. Isso garante que todos utilizem sempre a versão mais atualizada do aplicativo, acessando-o diretamente na rede em modo somente leitura.

A senha de edição é única e restrita aos responsáveis pela manutenção do aplicativo, garantindo maior segurança e controle sobre suas atualizações.

<h3 align="center">
  <img src="https://github.com/user-attachments/assets/e50c5097-a5fe-421b-a1c1-0a7e0624896f" width="600px" />
  <br><br>
</h3>






