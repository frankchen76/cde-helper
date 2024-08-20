import Debug from 'debug'

// Initialize debug logging module

export const info = Debug("cde-helper:info");
//info.log = console.log.bind(console)

export const error = Debug("cde-helper:error");
//error.log = console.log.bind(console)
