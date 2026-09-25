# Criador de Relatórios

Aplicativo desktop portátil para geração de Relatórios Fotográficos em Word (.docx), com identidade visual alinhada ao S.O.S. — Sistema de Ordens de Manutenção.

## Recursos
- Interface desktop para dados do relatório e seleção múltipla de fotos.
- Layouts de 2, 4, 6 e 8 fotos por página.
- Processamento sequencial de grandes volumes de imagens.
- Numeração automática das fotos.
- Documento A4 com cabeçalho repetido.
- Imagens preservando proporção e orientação EXIF.
- Sem dependência de caminhos absolutos de unidade.

## Gerar executável no Windows
```bat
py -m pip install -r requirements.txt
py -m PyInstaller --onefile --noconsole --name RelatorioFotografico relatorio_fotografico.py
```
