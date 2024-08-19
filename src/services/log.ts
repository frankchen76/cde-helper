import Debug from 'debug'

// Initialize debug logging module
export const log = Debug("podhelper");
log.log = console.log.bind(console)

export const info = Debug("podhelper:info");
info.log = console.log.bind(console)

export const err = Debug("podhelper:error");
err.log = console.log.bind(console)
