type HexFunction = (str: String) => Promise<Array<number>>
export const hex: HexFunction = async str =>
  str.split("").map(ch => ch.charCodeAt(0))

type SleepFunction = (ms: number) => Promise<void>
export const sleep: SleepFunction = ms => {
  return new Promise(resolve => setTimeout(resolve, ms))
}
