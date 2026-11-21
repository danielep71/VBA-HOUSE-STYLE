# Istruzioni per gli agenti

## Suggerimenti proattivi (non aspettare che te lo chieda)

1. **Merge in main**

   Dimmi quando è un buon momento per fare il merge in main, con una riga di
   motivo. Lo è se: il branch release/X.Y.Z ha le issue della milestone chiuse,
   i controlli sono verdi sull'head finale, la review (@codex) copre quell'head
   e il CHANGELOG è datato. Se manca qualcosa, dimmi cosa. Non fare il merge
   senza il mio via libera esplicito.

2. **/compact**

   Dimmi quando è un buon momento per lanciare /compact (dopo un merge, a fine
   issue, prima di cambiare argomento, o quando il contesto è pieno). Ogni volta
   proponi il comando completo, pronto da incollare, nel formato:

   ```text
   /compact Keep: <fatti da conservare> Drop: <dettagli da scartare>
   ```

   - Keep: stato del repo (branch, issue/PR aperte, versione), decisioni
     prese, convenzioni da rispettare, prossimo passo.
   - Drop: bozze, conflitti già risolti, output di comandi, esplorazioni
     chiuse.

   Basa il testo sulla sessione in corso: non copiare esempi di altri progetti.

3. **Modello prima di ogni nuova issue**

   Prima di implementarla, analizzala e suggerisci Sonnet, Opus o Fable, con una
   riga di motivo.

4. **Modello della sessione e dei sub agenti**

   - Sub agenti: scegli tu il modello più adatto, senza chiedere.
   - Sessione principale: se ritieni utile cambiare modello, chiedimelo con una
     domanda secca (modello proposto e motivo). Il cambio lo faccio io con
     /model.
