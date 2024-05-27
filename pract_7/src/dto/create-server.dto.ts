export class ConditionDto {
    type: string;
    value: any;
  }
  
  export class CreateServerDto {
    url: string;
    status: string;
    conditions: ConditionDto[];
  }
  