import { Routes } from '@angular/router';

export const routes: Routes = [
  {
    path: 'daybook',
    loadComponent: () =>
      import('./daybook/daybook').then(m => m.Daybook)
  },

  {
    path: '',
    redirectTo: 'daybook',
    pathMatch: 'full'
  }
];
