# Como Criar Documento Word Com Openxml

**Estou prestes a te mostrar como criar um documento word em c# com openxml, algo que me levou mais de 17 pomodoros de 50 min**😮😮😮

![pomodoros](https://user-images.githubusercontent.com/46006217/166085851-327c9335-4500-4b7e-b5d0-cd4ca7c7e99e.png)

Mas felizmente consegui criar e vou compartilhar com você, quero lembrar que não sou o dono da verdade e pode existir formas melhores de se fazer isso.

###### Dica e presente ao final 💡🎁
###### Qualquer contribuição é só fazer aquele PR esperto 😎
###### Caso queira mandar aquele pix maneiro de 1 centavo🪙🤑 👉 157.872.727-81 🔑

### Vamos ao passo a passo:

1. Baixe o Open XML SDK 2.5 Productivity Tool nesse link: [Open XML Productivity Tool][app openxml]

2. Abra o word e crie uma tabela, gráfico e qualquer outra coisa que você quiser, salve e feche.

3. Com a ferramenta baixada, abra e clique em Open File e escolha o arquivo.

![open file](https://user-images.githubusercontent.com/46006217/166085845-4dad6e56-12ae-44ee-83fb-0ceb051e944b.png)

4. Escolha qual parte do documento você vai querer refletir em c# ou deixe como vem para refletir o documento inteiro.

![get part](https://user-images.githubusercontent.com/46006217/166085848-eb9d533f-8bff-4480-851e-859c49fc1ad5.png)

5. clique em reflect code para gerar o código c# após escolher a parte do documento.

![reflect code](https://user-images.githubusercontent.com/46006217/166085847-ef6fa164-f574-4acd-afb7-7062015d07fd.png)

6. O código gerado ao lado em c# é exatamente a montagem do documento inteiro, copie o código ao lado faça as modificações e use como quiser.

![code source](https://user-images.githubusercontent.com/46006217/166085849-62100165-5f6f-4c58-b4be-59d275e8607f.png)

> 💡 O código é bem verboso e gigantesco, mas provavelmente suas mudanças vão ser feitas nas funções `GenerateMainDocumentPart1Content` e `GenerateChartPart1Content`.

**🎁 Vou deixar aqui tbm o código em que precisei fazer uma tabela com um gráfico em uma aplicação que gerava dados dinâmicos para um relatório, está todo comentado e com marcações para facilitar:**
 
##### Botão do código a baixo, só clicar😉👌
[![botão](https://user-images.githubusercontent.com/46006217/166087844-3328176b-986f-4fc8-99dc-0ff702ae53fb.png)](https://github.com/MaiconAvila/como-criar-word-com-openxml/blob/main/create-word-with-openxml.cs)

###### ❤️ Espero ter ajudado😁🤗

## License
GPL-3.0 License

[app openxml]: <https://github.com/OfficeDev/Open-XML-SDK/releases/download/v2.5/OpenXMLSDKV25.msi>
