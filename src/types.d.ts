export type JsonResult = {
    ranges: Array<Range>,
    total: number,
    username: string
}

export type Range = {
    seccion: number,
    score: number
}

export type OriginalUser = {
    email: string,
    fullName: Array<string>
}

export type UserRow = {
    firstSurname?: string,
    lastSurname?: string,
    firstname?: string,
    email?: string,
    results?: number
}