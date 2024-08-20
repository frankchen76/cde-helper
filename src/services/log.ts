import Debug from 'debug'

// Initialize debug logging module
export const log = Debug("cde-helper");
log.log = console.log.bind(console)

export const info = Debug("cde-helper:info");
info.log = console.log.bind(console)

export const err = Debug("cde-helper:error");
err.log = console.log.bind(console)
