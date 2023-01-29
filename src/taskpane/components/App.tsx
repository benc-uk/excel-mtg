import * as React from 'react'
import { DefaultButton, ProgressIndicator, PrimaryButton, Text } from '@fluentui/react'
import Splash from './Splash'
import { MagicCard } from '../../lib/types'
import { lookupCard } from '../../lib/api'

/* global Excel require */
/* eslint-disable office-addins/load-object-before-read */

export interface AppProps {
  title: string
  isOfficeInitialized: boolean
}

export interface AppState {
  inProgress?: boolean
  percentageDone?: number
}

export default class App extends React.Component<AppProps, AppState> {
  private cancelLookup: boolean

  constructor(props, context) {
    super(props, context)
    this.state = {}
    this.cancelLookup = false
  }

  componentDidMount() {
    this.setState({
      inProgress: false,
      percentageDone: 0,
    })
  }

  click = async () => {
    try {
      this.cancelLookup = false
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange()

        range.load('values')
        range.load('text')
        await context.sync()

        this.setState({
          inProgress: true,
        })

        for (let i = 0; i < range.text.length; i++) {
          const cardName = range.text[i][0] as string
          try {
            this.setState({
              percentageDone: i / range.text.length,
            })

            const card = await lookupCard(cardName)
            if (!card) {
              continue
            }

            if (this.cancelLookup) {
              break
            }

            range.getCell(i, 0).valuesAsJson = [[cardToEntity(card)]]
            await context.sync()
          } catch (error) {
            console.error(error)
          }
        }

        this.setState({
          inProgress: false,
          percentageDone: 0.0,
        })
      })
    } catch (error) {
      console.error(error)
    }
  }

  render() {
    const { title, isOfficeInitialized } = this.props

    if (!isOfficeInitialized) {
      return (
        <Splash
          title={title}
          logo={require('./../../../assets/mtg.png')}
          message="Please sideload your addin to see app body."
        />
      )
    }

    return (
      <main className="ms-welcome__main">
        <Text>Select one or more cells containing card names, then click the button</Text>
        <hr></hr>
        <PrimaryButton text="Lookup Card Names" onClick={this.click} disabled={this.state.inProgress} />

        {this.state.inProgress && (
          <div>
            <ProgressIndicator
              label="Looking up cards"
              description="Please wait..."
              percentComplete={this.state.percentageDone}
            />

            <DefaultButton
              text="Cancel"
              onClick={() => {
                this.cancelLookup = true
              }}
            />
          </div>
        )}
      </main>
    )
  }
}

function cardToEntity(card: MagicCard): Excel.EntityCellValue {
  const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: card.name,

    properties: {
      Name: {
        type: Excel.CellValueType.string,
        basicValue: card.name,
      },
      'Mana Cost': {
        type: Excel.CellValueType.string,
        basicValue: card.manaCost || '',
      },
      'Converted Cost': {
        type: Excel.CellValueType.double,
        basicValue: card.cmc,
      },
      'Type Line': {
        type: Excel.CellValueType.string,
        basicValue: card.type,
      },
      Text: {
        type: Excel.CellValueType.string,
        basicValue: card.text || '',
      },
      'Flavor Text': {
        type: Excel.CellValueType.string,
        basicValue: card.flavor || '',
      },
      Set: {
        type: Excel.CellValueType.string,
        basicValue: card.setName,
      },
      'Set Code': {
        type: Excel.CellValueType.string,
        basicValue: card.set,
      },
      Rarity: {
        type: Excel.CellValueType.string,
        basicValue: card.rarity,
      },
      'Multiverse ID': {
        type: Excel.CellValueType.string,
        basicValue: card.multiverseid || '',
      },
      Number: {
        type: Excel.CellValueType.string,
        basicValue: card.number,
      },
      Artist: {
        type: Excel.CellValueType.string,
        basicValue: card.artist,
      },
    },

    layouts: {
      card: {
        title: { property: 'Name' },
        mainImage: { property: 'Image' },
        subTitle: { property: 'Mana Cost' },
        sections: [
          {
            layout: 'List',
            title: 'Card Text',
            collapsed: true,
            collapsible: true,
            properties: ['Text', 'Flavor Text'],
          },
          {
            layout: 'List',
            title: 'Core Info',
            collapsed: true,
            collapsible: true,
            properties: ['Mana Cost', 'Converted Cost', 'Rarity'],
          },
          {
            layout: 'List',
            title: 'Type Details',
            collapsed: true,
            collapsible: true,
            properties: ['Type Line'],
          },
          {
            layout: 'List',
            title: 'Set Details',
            collapsed: true,
            collapsible: true,
            properties: ['Set', 'Set Code'],
          },
          {
            layout: 'List',
            title: 'Printing',
            collapsed: true,
            collapsible: true,
            properties: ['Number', 'Multiverse ID', 'Artist'],
          },
        ],
      },
    },
  }

  if (card.colors !== undefined) {
    const colors = []
    for (const color of card.colors) {
      colors.push({
        type: Excel.CellValueType.string,
        basicValue: color,
      })
    }
    entity.properties['Colors'] = {
      type: Excel.CellValueType.array,
      elements: [colors],
    }
    entity.layouts.card.sections[1].properties.push('Colors')
  }

  if (card.types !== undefined) {
    const types = []
    for (const type of card.types) {
      types.push({
        type: Excel.CellValueType.string,
        basicValue: type,
      })
    }
    entity.properties['Types'] = {
      type: Excel.CellValueType.array,
      elements: [types],
    }
    entity.layouts.card.sections[2].properties.push('Types')
  }

  if (card.supertypes !== undefined) {
    const supertypes = []
    for (const supertype of card.supertypes) {
      supertypes.push({
        type: Excel.CellValueType.string,
        basicValue: supertype,
      })
    }
    entity.properties['Supertypes'] = {
      type: Excel.CellValueType.array,
      elements: [supertypes],
    }
    entity.layouts.card.sections[2].properties.push('Supertypes')
  }

  if (card.subtypes !== undefined) {
    const subtypes = []
    for (const subtype of card.subtypes) {
      subtypes.push({
        type: Excel.CellValueType.string,
        basicValue: subtype,
      })
    }
    entity.properties['Subtypes'] = {
      type: Excel.CellValueType.array,
      elements: [subtypes],
    }
    entity.layouts.card.sections[2].properties.push('Subtypes')
  }

  if (card.printings !== undefined) {
    const printings = []
    for (const printing of card.printings) {
      printings.push({
        type: Excel.CellValueType.string,
        basicValue: printing,
      })
    }
    entity.properties['Printings'] = {
      type: Excel.CellValueType.array,
      elements: [printings],
    }
    entity.layouts.card.sections[3].properties.push('Printings')
  }

  if (card.imageUrl !== undefined) {
    entity.properties['Image'] = {
      type: Excel.CellValueType.webImage,
      address: card.imageUrl.replace('http://', 'https://'),
    }
  } else {
    entity.properties['Image'] = {
      type: Excel.CellValueType.webImage,
      address: 'https://static.wikia.nocookie.net/mtgsalvation_gamepedia/images/f/f8/Magic_card_back.jpg',
    }
  }

  if (card.power !== undefined) {
    entity.properties['Power'] = {
      type: Excel.CellValueType.string,
      basicValue: card.power,
    }
  }

  if (card.toughness !== undefined) {
    entity.properties['Toughness'] = {
      type: Excel.CellValueType.string,
      basicValue: card.toughness,
    }
  }

  if (card.toughness !== undefined && card.power !== undefined) {
    entity.layouts.card.sections.push({
      layout: 'List',
      title: 'Power/Toughness',
      collapsed: true,
      collapsible: true,
      properties: ['Power', 'Toughness'],
    })
  }

  return entity
}
