![Brazil Banner](https://iili.io/BGMGnEu.png)

# l2k_mloExtractor

## Important Credit

**Any donation link below is NOT for me.**

This support goes to **DexyFex**, the original creator of **CodeWalker**.  
Without CodeWalker and its source foundation, this tool would not exist in its current form.

- Patreon: https://www.patreon.com/dexyfex

**Base source:** DexyFex / CodeWalker  
**Workflow and support:** Codex

[PT-BR] Ferramenta para abrir arquivos `YTYP` ou `YTYP.XML` de MLO e exportar os props do interior com os arquivos e texturas relacionados, em uma estrutura mais pratica para workflow no Blender.

[EN] Tool designed to open MLO `YTYP` or `YTYP.XML` files and export interior props together with all related files and textures in a Blender-friendly structure.

---

## Portugues (PT-BR)

**Sempre atualize a pasta `mods` com props addon antes de abrir a aplicacao, para que o cache seja montado corretamente.**

**IMPORTANTE:** copie todos os arquivos addon para a pasta `mods` e crie um `RPF` com um nome unico dentro de `mods`.  
Isso ajuda a ferramenta a resolver corretamente os hashes e localizar os props addon, evitando falhas e recursos ausentes na hora de exportar para o Blender. Se usar o campo `Addon RPF`, informe esse nome unico para evitar colisao com outros pacotes.

### Como usar

1. Faca backup do seu mod antes de qualquer teste.
2. Rode `Blender MLO Extractor.exe`.
3. Para props addon, crie um `RPF` com nome unico dentro da pasta `mods` e copie o conteudo do addon para dentro dele.
4. Espere a ferramenta terminar o carregamento inicial dos arquivos do GTA.
5. Abra o `YTYP` onde esta o seu MLO. Nao e necessario converter para XML, mas `*.ytyp.xml` tambem e suportado.
6. A ferramenta vai criar uma pasta ao lado do arquivo original no formato `Mlo Extracted - <nome>`.
7. Dentro dela, o XML principal sai com o nome do arquivo original, por exemplo `meu_interior.ytyp.xml`.
8. Os modelos e texturas relacionados sao exportados para a pasta `Drawable`.
9. No Blender, o fluxo recomendado e importar primeiro o conteudo de `Drawable` e depois importar o `arquivo.ytyp.xml` para aplicar a estrutura e o posicionamento do interior.

### Observacoes

- Em alguns casos a primeira carga pode demorar um pouco, principalmente quando existem props addon.
- A ferramenta tenta localizar modelos e texturas externas, inclusive dentro de `mods`.
- Texturas compartilhadas agora sao cacheadas por YTD durante a exportacao para evitar reabrir o mesmo dictionary varias vezes.
- O resultado da exportacao e salvo ao lado do arquivo original para facilitar a organizacao.
- Se aparecer `prop archetypes could not be resolved` ou `prop resources were not found`, normalmente isso significa que parte do conteudo referenciado nao estava disponivel no GTA, `mods` ou RPF carregado durante a extracao.

### Tutorial

- Video tutorial placeholder: https://www.youtube.com/watch?v=dQw4w9WgXcQ

---

## English (EN)

**Always update the `mods` folder with addon props before opening the application so the cache can be built correctly.**

**IMPORTANT:** copy all addon files into the `mods` folder and create an `RPF` with a unique name inside `mods`.  
This helps the tool resolve hashes correctly and locate addon props, avoiding missing resources and export problems when sending the result to Blender. If you use the `Addon RPF` field, enter that same unique name to avoid collisions with other packages.

### How to use

1. Back up your mod before testing anything.
2. Run `Blender MLO Extractor.exe`.
3. For addon props, create an `RPF` with a unique name inside the `mods` folder and copy the addon content into it.
4. Wait until the tool finishes loading the GTA files.
5. Open the `YTYP` that contains your MLO. You do not need to convert it to XML, although `*.ytyp.xml` is also supported.
6. The tool will create a folder next to the original file using the format `Mlo Extracted - <name>`.
7. Inside it, the main XML keeps the original file name, for example `my_interior.ytyp.xml`.
8. Related models and textures are exported to the `Drawable` folder.
9. In Blender, the recommended workflow is to import the contents of `Drawable` first and then import the root `file.ytyp.xml` to apply the MLO structure and placement.

### Notes

- In some cases the initial load can take a while, especially when addon props are involved.
- The tool attempts to locate external models and textures, including files inside `mods`.
- Shared textures are now cached per YTD during export to avoid reopening the same dictionary repeatedly.
- Export output is saved next to the original file for easier organization.
- If you see `prop archetypes could not be resolved` or `prop resources were not found`, it usually means some referenced content was not available in the loaded GTA, `mods`, or RPF data during extraction.

### Tutorial

- Video tutorial placeholder: https://www.youtube.com/watch?v=dQw4w9WgXcQ

---

## Local Build

This package includes a standalone build script so it can be compiled without depending on the full CodeWalker solution.

```powershell
powershell -ExecutionPolicy Bypass -File .\build_release.ps1
```

The executable will be generated at:

```text
build\Release\Blender MLO Extractor.exe
```
