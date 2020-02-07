import { TestBed } from '@angular/core/testing';

import { NewexcelService } from './newexcel.service';

describe('NewexcelService', () => {
  beforeEach(() => TestBed.configureTestingModule({}));

  it('should be created', () => {
    const service: NewexcelService = TestBed.get(NewexcelService);
    expect(service).toBeTruthy();
  });
});
