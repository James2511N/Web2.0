import { ApiProperty } from '@nestjs/swagger'
import {
    IsNotEmpty,
    IsNumber,
    IsObject,
    IsString,
    ValidateNested,
} from 'class-validator'
import { Type } from 'class-transformer'

class LocationDto {
    @ApiProperty({ example: 2.5052 })
    @IsNumber()
    @IsNotEmpty()
    longitude: number

    @ApiProperty({ example: 7.5628 })
    @IsNumber()
    @IsNotEmpty()
    latitude: number
}

export class AddressesDto {
    @ApiProperty({ type: String })
    @IsString()
    @IsNotEmpty()
    _id: string

    @ApiProperty({ type: String })
    @IsString()
    @IsNotEmpty()
    name: string

    @ApiProperty({ type: LocationDto })
    @IsObject()
    @ValidateNested()
    @Type(() => LocationDto)
    location: LocationDto
}
