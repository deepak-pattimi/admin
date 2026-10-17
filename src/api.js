import axios from 'axios';

export const API_URL = (import.meta.env.VITE_API_URL || 'http://localhost:3016').replace(/\/+$/, '');
axios.defaults.timeout = 15000;
