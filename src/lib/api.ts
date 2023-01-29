import { MagicCard } from './types'

const apiBase = 'https://api.magicthegathering.io/v1'

// ============================================================
// Lookup a Magic card by name
// ============================================================
export async function lookupCard(name: string): Promise<MagicCard> {
  if (!name) {
    return null
  }

  const resp = await fetch(`${apiBase}/cards?name=${name}`)

  if (!resp.ok) {
    throw new Error(resp.statusText)
  }

  const allCards = await resp.json()

  for (const card of allCards.cards) {
    if (card.name === name) {
      console.log(`### Located card: ${card.name} (${card.setName})`)
      console.log(card)
      return card as MagicCard
    }
  }

  return null
}
