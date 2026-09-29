import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Daybook } from './daybook';

describe('Daybook', () => {
  let component: Daybook;
  let fixture: ComponentFixture<Daybook>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Daybook],
    }).compileComponents();

    fixture = TestBed.createComponent(Daybook);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
