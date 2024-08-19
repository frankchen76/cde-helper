import Debug from 'debug'

// Initialize debug logging module

export const info = Debug("podhelper:info");
//info.log = console.log.bind(console)

export const error = Debug("podhelper:error");
//error.log = console.log.bind(console)
