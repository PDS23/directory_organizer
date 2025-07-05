#  Organizador de Arquivos por Planilha

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

Esta é uma ferramenta com interface gráfica (GUI) desenvolvida em Python e PyQt5 para automatizar a organização de arquivos. O script move arquivos de uma pasta de origem para uma estrutura de pastas de destino hierárquica, baseando-se em uma planilha de mapeamento (Excel ou CSV).

<!-- ## 📸 Screenshot

(Adicione aqui um screenshot da aplicação em funcionamento para que todos possam ver a interface)
![Exemplo da Interface](https://via.placeholder.com/700x500.png?text=Interface+do+Organizador+de+Arquivos) -->

## ✨ Funcionalidades Principais

-   **Interface Gráfica Amigável**: Facilita a seleção de pastas e da planilha, sem necessidade de alterar o código.
-   **Suporte a Múltiplos Formatos**: Lê planilhas `.xlsx` e `.csv`.
-   **Filtro por Categoria**: Permite selecionar quais categorias da planilha devem ser processadas.
-   **Lógica de Correspondência Inteligente**: Associa arquivos a dados da planilha usando um prefixo no nome do arquivo.
-   **Seleção da Correspondência Mais Recente**: Se houver múltiplos registros para o mesmo arquivo, o script utiliza o **último** encontrado na planilha (o mais recente).
-   **Criação Automática de Diretórios**: Gera a estrutura de pastas `Unidade/DATA/OP/OS` no destino.
-   **Log de Processamento**: Exibe em tempo real o status de cada arquivo processado.
-   **Tratamento de Arquivos "Órfãos"**: Arquivos que não encontram correspondência na planilha são movidos para uma pasta `Sem_Correspondencia`.

## 📝 Estrutura da Planilha

Para que o script funcione corretamente, a sua planilha **deve** conter as seguintes colunas:

| Nome da Coluna         | Descrição                                                                                             | Exemplo          |
| ---------------------- | ------------------------------------------------------------------------------------------------------- | ---------------- |
| `Bloco`                | O identificador que será comparado com o prefixo do nome do arquivo.                                    | `12345`          |
| `Unidade`              | O nome da unidade, que será usado para criar a primeira pasta de destino.                               | `Unidade_SP`     |
| `Número da Operação`   | O número da OP, usado para a subpasta.                                                                  | `OP-987`         |
| `Número da OS`         | O número da OS, usado para a última subpasta da hierarquia.                                             | `OS-654`         |
| `Categoria`            | Uma categoria para agrupar os itens. Usado no filtro inicial.                                           | `Tipo A`         |

## 📂 Nomenclatura dos Arquivos

Os arquivos na pasta de origem precisam seguir uma convenção de nome para que a correspondência funcione. O prefixo do arquivo (a parte antes do primeiro `_`) deve corresponder a um valor na coluna `Bloco` da planilha.

**Formato esperado:** `BLOCO_RestoDoNome.ext`

**Exemplo:**
-   Nome do arquivo: `12345_relatorio_final.pdf`
-   O script extrairá o prefixo: `12345`
-   Em seguida, buscará por `12345` na coluna `Bloco` da planilha.

## 🚀 Tecnologias Utilizadas

-   [Python 3](https://www.python.org/)
-   [PyQt5](https://riverbankcomputing.com/software/pyqt/) - Para a interface gráfica.
-   [Pandas](https://pandas.pydata.org/) - Para leitura e manipulação da planilha.
-   [Openpyxl](https://openpyxl.readthedocs.io/en/stable/) - Requerido pelo Pandas para ler arquivos `.xlsx`.

## ✅ Pré-requisitos

-   Python 3.x instalado.
