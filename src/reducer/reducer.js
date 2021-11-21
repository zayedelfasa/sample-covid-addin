import { IS_INITIALIZED, IS_LOADING, IS_DONE, IS_ERROR } from '../reducer/status';

const Reducer = (state, action) => {
    switch (action.type) {
        case IS_INITIALIZED:
            return { ...state, data: action.data == [] ? [] : {}, error: null, loadState: IS_INITIALIZED };
        case IS_LOADING:
            return { ...state, data: action.data == [] ? [] : {}, error: null, loadState: IS_LOADING};
        case IS_DONE:
            return { ...state, data: action.data , error: null, loadState: IS_DONE };
        case IS_ERROR:
            return { ...state, data: [], error: action.error, loadState: IS_ERROR };
        default:
            throw new Error();
    }
}

export default Reducer;